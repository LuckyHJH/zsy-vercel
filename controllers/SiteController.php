<?php

namespace app\controllers;

use Yii;
use yii\web\Controller;
use yii\web\Response;
use yii\web\UploadedFile;
use app\models\UploadForm;

class SiteController extends Controller
{
    /**
     * Displays homepage.
     *
     * @return string
     */
    public function actionIndex()
    {
        $model = new UploadForm();
        if (!Yii::$app->request->isPost) {
            return $this->render('index', ['model' => $model]);
        } else {
            $model->file = UploadedFile::getInstance($model, 'file');
            ini_set('memory_limit','-1');
            //set_time_limit(0);
            $file = $model->file->tempName;
            //$file = '../web/1.xlsm';

            $objPHPExcelOutput = new \PHPExcel();
            $objPHPExcelOutput->setActiveSheetIndex(0);
            $objActSheetOutput = $objPHPExcelOutput->getActiveSheet();
            $objActSheetOutput->setTitle('Sheet1');
            $objReader = \PHPExcel_IOFactory::createReader('Excel2007');//use excel2007 for 2007 format
            $objReader = $objReader->load($file);

            $NUM_COL = 0;
            $DATE_COL = 1;
            $SHOP_COL = 2;
            $NAME_COL = 6;
            $AMOUNT_COL = 12;
            $PRICE_COL = 13;
            $TOTAL_COL = 14;

            echo '<table>';
            $iSheetCount = $objReader->getSheetCount();
            $aData = [];
            for ($i = 0;$i < $iSheetCount;$i++) {
                $curSheetInput = $objReader->getSheet($i);
                $iRow = 6;
                while (true) {
                    $oRow = [];
                    $oRow['num'] = $curSheetInput->getCellByColumnAndRow($NUM_COL, $iRow)->getValue();
                    $oRow['date'] = $curSheetInput->getCellByColumnAndRow($DATE_COL, $iRow)->getValue();
                    if (empty($oRow['date'])) {
                        break;
                    }
                    $oRow['shop'] = $curSheetInput->getCellByColumnAndRow($SHOP_COL, $iRow)->getValue();
                    $oRow['name'] = $curSheetInput->getCellByColumnAndRow($NAME_COL, $iRow)->getValue();
                    $oRow['amount'] = $curSheetInput->getCellByColumnAndRow($AMOUNT_COL, $iRow)->getValue();
                    $oRow['price'] = $curSheetInput->getCellByColumnAndRow($PRICE_COL, $iRow)->getValue();
                    $oRow['total'] = $curSheetInput->getCellByColumnAndRow($TOTAL_COL, $iRow)->getValue();
                    $aData[] = $oRow;
                    $iRow++;
                    echo '<tr>';
                    echo '<td>'.$oRow['num'].'</td>';
                    echo '<td>'.$oRow['date'].'</td>';
                    echo '<td>'.$oRow['shop'].'</td>';
                    echo '<td>'.$oRow['name'].'</td>';
                    echo '<td>'.$oRow['amount'].'</td>';
                    echo '<td>'.$oRow['price'].'</td>';
                    echo '<td>'.$oRow['total'].'</td>';
                    echo '</tr>';
                }
            }
            echo '</table>';
            //yiiwebResponse::sendFile();
        }
    }

    /**
     * Login action.
     *
     * @return Response|string
     */
    public function actionLogin()
    {
        if (!Yii::$app->user->isGuest) {
            return $this->goHome();
        }

        $model = new LoginForm();
        if ($model->load(Yii::$app->request->post()) && $model->login()) {
            return $this->goBack();
        }
        return $this->render('login', [
            'model' => $model,
        ]);
    }

    /**
     * Logout action.
     *
     * @return Response
     */
    public function actionLogout()
    {
        Yii::$app->user->logout();

        return $this->goHome();
    }

    /**
     * Displays contact page.
     *
     * @return Response|string
     */
    public function actionContact()
    {
        $model = new ContactForm();
        if ($model->load(Yii::$app->request->post()) && $model->contact(Yii::$app->params['adminEmail'])) {
            Yii::$app->session->setFlash('contactFormSubmitted');

            return $this->refresh();
        }
        return $this->render('contact', [
            'model' => $model,
        ]);
    }

    /**
     * Displays about page.
     *
     * @return string
     */
    public function actionAbout()
    {
        return $this->render('about');
    }
}
