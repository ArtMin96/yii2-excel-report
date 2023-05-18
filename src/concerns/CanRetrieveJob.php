<?php

namespace minasyans\excelreport\concerns;

use Yii;
use yii\db\Exception;

trait CanRetrieveJob
{
    /**
     * @throws Exception
     */
    public function retrieveJob(int $id)
    {
        return Yii::$app
            ->db
            ->createCommand(
                "SELECT * FROM {{%queue}} WHERE id = :id",
                [':id' => $id]
            )
            ->queryOne();
    }
}