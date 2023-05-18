<?php

namespace minasyans\excelreport\controllers;

use minasyans\excelreport\concerns\CanRetrieveJob;
use Yii;
use yii\web\Controller;
use yii\web\NotFoundHttpException;

class ReportController extends Controller {
    use CanRetrieveJob;

    public function actionQueue() {
        Yii::$app->response->format = \yii\web\Response::FORMAT_JSON;
        $jobId = $_POST['id'];
        $data = [];
        $ready = false;
        $progress = Yii::$app->queue->getProgress($jobId);
        $job = $this->retrieveJob($jobId);

        if (Yii::$app->session->has('excel-report-progress-' . $_POST['name'])) {
            $data = unserialize(Yii::$app->session->get('excel-report-progress-' . $_POST['name']));
        }

        if (($progress[0] == $progress[1]) || ($progress[0] > 0 && ! $job)) {
            $ready = true;
        }

        return  [
            'progress' => $progress,
            'info' => $data,
            'ready' => $ready,
        ];
    }

    public function actionDownload() {
        if (Yii::$app->session->has('excel-report-progress-' . $_GET['name'])) {
            $data = unserialize(Yii::$app->session->get('excel-report-progress-' . $_GET['name']));
            $file = Yii::$app->basePath . '/runtime/export/' . $data['fileName'] . '.xlsx';

            if (file_exists($file)) {
                if (filesize($file) == 0) {
                    throw new NotFoundHttpException('The file is finished forming. Only a few seconds left... Try clicking the link again');
                    return false;
                } else {
                    Yii::$app->session->remove('excel-report-progress-' . $_GET['name']);
                    ob_clean();
                    return \Yii::$app->response->sendFile($file, null, ['mimeType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']);
                }
            } else {
                Yii::$app->session->remove('excel-report-progress-' . $_GET['name']);
                throw new NotFoundHttpException('This file does not exist');
            }
        } else {
            throw new NotFoundHttpException('This file does not exist');
        }
    }

    public function actionReset() {
        $jobId = $_POST['id'];
        Yii::$app->queue->setManuallyProgress($jobId, 1, 1);
        Yii::$app->session->remove('excel-report-progress-' . $_POST['name']);
    }
}
