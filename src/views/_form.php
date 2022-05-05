<?php

use yii\helpers\Html;

echo Html::beginForm('', 'post', $options);
echo Html::hiddenInput('excelReportAction', 1);
echo Html::submitButton(Yii::t('minasyans', 'Create Excel'), ['class' => 'btn btn-outline-primary ml-2']);
echo Html::endForm();