<?php

namespace minasyans\excelreport;

use minasyans\excelreport\ExcelReportHelper;
use Yii;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Common\Exception\UnsupportedTypeException;
use Box\Spout\Common\Type;
use Box\Spout\Writer\Style\BorderBuilder;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\Style\Border;
use Box\Spout\Writer\Style\Color;
use yii\base\InvalidConfigException;
use box\spout;
use yii\data\ArrayDataProvider;
use yii\queue\Queue;

/**
 * ExcelReportModel handles the generation of Excel reports from Yii2 data providers.
 *
 * This class provides high-performance Excel report generation using the Box/Spout library,
 * supporting both ArrayDataProvider and ActiveDataProvider as data sources.
 *
 * @property-read Writer $_writer The Box/Spout writer instance
 * @property-read mixed $_provider The data provider (ArrayDataProvider or ActiveDataProvider)
 * @property-read array $_columns The processed column configurations
 *
 * @since 1.0.0
 */
class ExcelReportModel {
    /** @var mixed Data provider instance */
    private $_provider;

    /** @var array Processed column configurations */
    private $_columns;

    /** @var array Cached column accessors and metadata */
    private $_columnCache = [];

    /** @var \yii\i18n\Formatter Yii formatter instance */
    private $_formatter;

    /** @var Writer Box/Spout writer instance */
    private $_writer;

    /** @var array Default style for data rows */
    private $_defaultRowStyle;

    /** @var array Style for header row */
    private $_headerStyle;

    /** @var int Number of records to process in each batch */
    private const BATCH_SIZE = 1000;

    /** @var int How often to report progress to queue */
    private const PROGRESS_REPORT_FREQUENCY = 1000;

    /** @var string Output filename */
    public $filename;

    /** @var bool Whether to strip HTML from cell values */
    public $stripHtml = true;

    /** @var string Output folder path */
    public $folder = '@app/runtime/export';

    /** @var Queue Queue instance for progress reporting */
    protected $queue;

    /**
     * Initialize the Excel report generator.
     *
     * @param string $columns Base64 encoded serialized array of gridview columns
     * @param Queue $queue Queue instance for progress reporting
     * @param string $fileName Output filename without extension
     * @param string $dataProvider Base64 encoded serialized data provider
     * @param array $config Additional configuration options
     *
     * @throws \Exception If data provider or columns cannot be unserialized
     */
    public function __construct($columns, $queue, $fileName, $dataProvider, array $config = [])
    {
        $this->_provider = ExcelReportHelper::reverseClosureDetect(unserialize(base64_decode($dataProvider)));
        $this->_columns = $this->optimizeColumns(ExcelReportHelper::reverseClosureDetect(unserialize(base64_decode($columns))));
        $this->queue = $queue;
        $this->filename = $fileName;
        $this->_formatter = Yii::$app->formatter;

        $this->initializeStyles();
        $this->prepareColumnAccessors();
    }

    /**
     * Filter and optimize column configurations.
     *
     * @param array $columns Raw column configurations
     * @return array Filtered and optimized columns
     */
    protected function optimizeColumns($columns)
    {
        return array_values(array_filter($columns, function($column) {
            return empty($column['hiddenFromExport']) &&
                (!isset($column['class']) || !in_array($column['class'], [
                        'yii\\grid\\ActionColumn',
                        'yii\\grid\\SerialColumn',
                        'yii\\grid\\CheckboxColumn'
                    ]));
        }));
    }

    /**
     * Initialize header and row styles.
     *
     * @return void
     */
    protected function initializeStyles()
    {
        // Create header style
        $headerBorder = (new BorderBuilder())
            ->setBorderBottom(Color::BLACK, Border::WIDTH_MEDIUM, Border::STYLE_SOLID)
            ->build();

        $this->_headerStyle = (new StyleBuilder())
            ->setFontBold()
            ->setBackgroundColor('FFE5E5E5')
            ->setBorder($headerBorder)
            ->build();

        // Create default row style
        $rowBorder = (new BorderBuilder())
            ->setBorderBottom(Color::BLACK, Border::WIDTH_MEDIUM, Border::STYLE_SOLID)
            ->build();

        $this->_defaultRowStyle = (new StyleBuilder())
            ->setBorder($rowBorder)
            ->build();
    }

    /**
     * Prepare column accessors for efficient data access.
     *
     * @return void
     */
    protected function prepareColumnAccessors()
    {
        foreach ($this->_columns as $index => $column) {
            $attribute = $column['attribute'] ?? null;
            $format = $column['format'] ?? null;
            $label = $column['label'] ?? ($column['header'] ?? '#');

            if (isset($column['value']) && is_callable($column['value'])) {
                $accessor = $column['value'];  // Use callable directly
                $isComputed = true; // Flag for computed value
            } elseif (is_string($attribute)) {
                $accessor = $this->compileStringAccessor($attribute);
                $isComputed = false;
            } elseif (is_callable($attribute)) {
                $accessor = $attribute;
                $isComputed = true;
            } else {
                $accessor = fn() => null;
                $isComputed = false;
            }

            $this->_columnCache[$index] = [
                'accessor' => $accessor,
                'format' => $format,
                'label' => $label,
                'isComputed' => $isComputed, // Store the flag
            ];
        }
    }

    /**
     * Compiles an accessor for string attributes (optimized).
     * @param string $attribute
     * @return callable
     */
    protected function compileStringAccessor(string $attribute): callable
    {
        $path = explode('.', $attribute);

        if (count($path) === 1) { // Direct attribute access
            return function ($model) use ($attribute) {
                return is_object($model) ? $model->$attribute : ($model[$attribute] ?? null);
            };
        } else { // Nested attribute access
            return function ($model) use ($path) {
                $current = $model;
                foreach ($path as $key) {
                    if (is_object($current)) {
                        $current = $current->$key;
                    } elseif (is_array($current)) {
                        $current = $current[$key] ?? null;
                    } else {
                        return null; // Early return if not object/array
                    }
                }
                return $current;
            };
        }
    }

    /**
     * Start the Excel report generation process.
     *
     * @throws InvalidConfigException|IOException|UnsupportedTypeException
     * @return void
     */
    public function start()
    {
        try {
            $this->initWriter();
            $this->writeHeader();
            $totalCount = $this->writeBody();
            $this->_writer->close();
            $this->queue->setProgress($totalCount, $totalCount);
        } catch (\Exception $e) {
            Yii::error($e->getMessage());
            throw $e;
        } finally {
            $this->cleanup();
        }
    }

    /**
     * Initialize the Box/Spout writer.
     *
     * @throws InvalidConfigException If the output directory cannot be created
     * @throws IOException If the file cannot be opened for writing
     * @throws UnsupportedTypeException If XLSX format is not supported
     * @return void
     */
    protected function initWriter()
    {
        $folder = trim(Yii::getAlias($this->folder));
        if (!file_exists($folder) && !mkdir($folder, 0777, true)) {
            throw new InvalidConfigException("Cannot create directory: $folder");
        }

        $file = rtrim($folder, '/\\') . DIRECTORY_SEPARATOR . $this->filename . '.xlsx';

        $this->_writer = WriterFactory::create(Type::XLSX);
        $this->_writer->setShouldUseInlineStrings(false);
        $this->_writer->openToFile($file);

        $this->_writer->getCurrentSheet()->setName(Yii::t('minasyans', 'Report'));
    }

    /**
     * Write the header row to the Excel file.
     *
     * @return void
     */
    protected function writeHeader()
    {
        if (empty($this->_columnCache)) {
            return;
        }

        $headers = array_column($this->_columnCache, 'label');
        $this->_writer->addRowWithStyle($headers, $this->_headerStyle);
    }

    /**
     * Write the data rows to the Excel file.
     *
     * @return int Total number of rows processed
     */
    protected function writeBody()
    {
        $totalCount = $this->_provider->getTotalCount();
        $processedCount = 0;
        $rows = [];

        $dataIterator = $this->_provider instanceof ArrayDataProvider
            ? [$this->_provider->allModels]
            : $this->_provider->query->batch(self::BATCH_SIZE);

        foreach ($dataIterator as $batch) {
            $models = $this->_provider instanceof ArrayDataProvider ? [$batch] : $batch;

            foreach ($models as $model) {
                $rows[] = $this->generateRow($model);
                $processedCount++;

                if (count($rows) >= self::BATCH_SIZE) {
                    $this->writeRows($rows);
                    $rows = [];
                }

                if ($processedCount % self::PROGRESS_REPORT_FREQUENCY === 0) {
                    $this->queue->setProgress($processedCount, $totalCount);
                }
            }
        }

        if (!empty($rows)) {
            $this->writeRows($rows);
        }

        return $totalCount;
    }

    /**
     * Generate a single data row.
     *
     * @param mixed $model The data model
     * @return array Row data
     */
    protected function generateRow($model)
    {
        $row = [];

        // Pre-compute values for closures only ONCE per row
        $computedValues = [];
        foreach ($this->_columnCache as $columnIndex => $columnData) {
            if ($columnData['isComputed']) {  // Check the flag directly
                $computedValues[$columnIndex] = $columnData['accessor']($model);
            }
        }

        foreach ($this->_columnCache as $columnIndex => $columnData) {
            $value = $columnData['isComputed'] // Use pre-computed value if available
                ? $computedValues[$columnIndex]
                : $columnData['accessor']($model); // Otherwise, call the accessor

            $row[] = $columnData['format']
                ? $this->_formatter->format($value, $columnData['format'])
                : $value;
        }
        return $row;
    }

    /**
     * Write a batch of rows to the Excel file.
     *
     * @param array &$rows Array of row data
     * @return void
     */
    protected function writeRows(array &$rows)
    {
        if ($this->stripHtml) {
            array_walk_recursive($rows, function(&$item) {
                $item = is_string($item) ? strip_tags($item) : $item;
            });
        }

        foreach ($rows as $row) {
            $this->_writer->addRowWithStyle($row, $this->_defaultRowStyle);
        }
    }

    /**
     * Clean up resources after export completion.
     *
     * @return void
     */
    protected function cleanup()
    {
        $this->_writer = null;
        $this->_provider = null;
        $this->_columnCache = [];
        gc_collect_cycles();
    }
}