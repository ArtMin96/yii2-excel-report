<?php

namespace minasyans\excelreport;

use Yii;
use yii\base\BootstrapInterface;

class Bootstrap implements BootstrapInterface{
    public function bootstrap($app)
    {
        $app->setModule('excelreport', 'minasyans\excelreport\Module');
        Yii::$app->getModule('excelreport')->registerTranslations();
    }
}