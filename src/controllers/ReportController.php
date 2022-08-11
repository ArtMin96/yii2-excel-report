<?php

namespace minasyans\excelreport\controllers;

use Yii;
use yii\web\Controller;
use yii\web\NotFoundHttpException;

class ReportController extends Controller {
    public function actionQueue(Request $request) {
        Yii::$app->response->format = \yii\web\Response::FORMAT_JSON;
        $jobId = $request->post('id');
        $data = [];

        if (Yii::$app->session->has('excel-report-progress-' . $request->post('name'))) {
            $data = unserialize(Yii::$app->session->get('excel-report-progress-' . $request->post('name')));
        }

        return  [
            'progress' => Yii::$app->queue->getProgress($jobId),
            'info' => $data,
        ];
    }

    public function actionDownload(Request $request) {
        if (Yii::$app->session->has('excel-report-progress-' . $request->get('name'))) {
            $data = unserialize(Yii::$app->session->get('excel-report-progress-' . $request->get('name')));
            $file = Yii::$app->basePath . '/runtime/export/' . $data['fileName'] . '.xlsx';

            if (file_exists($file)) {
                if (filesize($file) == 0) {
                    throw new NotFoundHttpException('Файл заканчивает формирование. Осталось всего несколько секунд... Попробуйте нажать на ссылку еще раз ');
                    return false;
                } else {
                    Yii::$app->session->remove('excel-report-progress');
                    ob_clean();
                    return \Yii::$app->response->sendFile($file, null, ['mimeType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']);
                }
            } else {
                Yii::$app->session->remove('excel-report-progress');
                throw new NotFoundHttpException('Такого файла не существует ');
            }
        } else {
            throw new NotFoundHttpException('Такого файла не существует ');
        }
    }

    public function actionReset(Request $request) {
        $jobId = $request->post('id');
        Yii::$app->queue->setManuallyProgress($jobId, 1, 1);
        Yii::$app->session->remove('excel-report-progress-' . $request->post('name'));
    }
}
