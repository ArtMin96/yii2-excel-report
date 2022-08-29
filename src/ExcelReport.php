<?php

namespace minasyans\excelreport;

use Yii;
use yii\base\Widget;
use minasyans\excelreport\ExcelReportHelper;

class ExcelReport extends Widget
{
    public $name = 'excel_report_form';
    public $columns;
    public $stripHtml;
    public $dataProvider;

    private $formOptions;
    private $id = "excelReportFrom";

    public function init()
    {
        parent::init();
        $this->formOptions = ['id' => $this->id,];
    }

    public function run()
    {
        if (isset($_POST['excelReportAction']) || Yii::$app->session->has('excel-report-progress-' . $this->name)) {
            if (Yii::$app->session->has('excel-report-progress-' . $this->name)) {
                $data = unserialize(Yii::$app->session->get('excel-report-progress-' . $this->name));
                $fileName = $data['fileName'];
                $id = $data['queueid'];
            } else {
                $this->columns = base64_encode(serialize(ExcelReportHelper::closureDetect($this->columns)));
                $this->dataProvider = base64_encode(serialize(ExcelReportHelper::closureDetect($this->dataProvider)));
                $fileName = Yii::$app->security->generateRandomString();
                $id = Yii::$app->queue->push(new ExcelReportQueue([
                    'columns' => $this->columns,
                    'stripHtml' => $this->stripHtml,
                    'fileName' => $fileName,
                    'dataProvider' => $this->dataProvider,
                ]));

                Yii::$app->session->set('excel-report-progress-' . $this->name, serialize(['fileName' => $fileName, 'queueid' => $id]));
            }

            return $this->render('_progress', [
                'options' => $this->formOptions,
                'queueId' => $id,
                'name' => $this->name,
                'file' => Yii::$app->basePath . '/runtime/export/' . $fileName . '.xlsx',
            ]);
        } else {
            return $this->render('_form', ['options' => $this->formOptions]);
        }
    }
}