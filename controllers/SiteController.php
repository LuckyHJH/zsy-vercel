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
        if (!Yii::$app->request->isPost) {
            $model = new UploadForm();
            return $this->render('index', ['model' => $model]);
        } else {
            try {
                ini_set('memory_limit','-1');
                //set_time_limit(0);
                $type = Yii::$app->request->post('type');
                return $this->$type();
            } catch (\Exception $e) {
                $content = '发生错误,请联系帅哥LuckyHJH';
                Yii::error(array(
                    'message'=>$e->getMessage(),
                    'trace'=>$e->getTrace()[0],
                ));
                return $content;
            }
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

    //退货（表1）
    private function tuihuo()
    {
        $NUM_COL = 0;
        $DATE_COL = 1;
        $SHOP_COL = 2;
        $NAME_COL = 6;
        $AMOUNT_COL = 12;
        $PRICE_COL = 13;
        $TOTAL_COL = 14;
        $START_ROW = 6;

        $model = new UploadForm();
        $model->file = UploadedFile::getInstance($model, 'file');
        $file = $model->file->tempName;
        //$file = '../web/1.xlsm';
        $objReader = \PHPExcel_IOFactory::createReader('Excel2007');//use excel2007 for 2007 format
        $objReader = $objReader->load($file);

        $html = '<table>';
        $iSheetCount = $objReader->getSheetCount();
        $aData = [];
        for ($i = 0;$i < $iSheetCount;$i++) {
            $curSheetInput = $objReader->getSheet($i);
            $iRow = $START_ROW;
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
                $html .= '<tr>';
                $html .= '<td>'.$oRow['num'].'</td>';
                $html .= '<td>'.$oRow['date'].'</td>';
                $html .= '<td>'.$oRow['shop'].'</td>';
                $html .= '<td>'.$oRow['name'].'</td>';
                $html .= '<td>'.$oRow['amount'].'</td>';
                $html .= '<td>'.$oRow['price'].'</td>';
                $html .= '<td>'.$oRow['total'].'</td>';
                $html .= '</tr>';
            }
        }
        $html .= '</table>';

        return $html;
    }

    //TODO 表2
    private function zhuangxiang()
    {
    }

    //店别（表3）
    private function dianbie()
    {
        $C_COL = 2;

        $model = new UploadForm();
        $model->file = UploadedFile::getInstance($model, 'file');
        $file = $model->file->tempName;
        //$file = '../web/3.xlsm';
        $objReader = \PHPExcel_IOFactory::createReader('Excel2007');//use excel2007 for 2007 format
        $objReader = $objReader->load($file);
        $objOutput = new \PHPExcel();
        $objOutput->getProperties()->setCreator('LuckyHJH');

        $iSheetCount = $objReader->getSheetCount();
        for ($i = 0;$i < $iSheetCount;$i++) {
            $write_row = 8;
            $curSheetInput = $objReader->getSheet($i);
            $title = $curSheetInput->getTitle();
            if ($i != 0) $objOutput->createSheet();
            $objActSheetOutput = $objOutput->getSheet($i);
            $objActSheetOutput->setTitle($title);
            $objActSheetOutput->SetCellValue("A1", "广州市志祥服装有限责任公司送货单");
            $objActSheetOutput->SetCellValue("A2", "地址：广州市西槎路天枢街3号3楼");
            $objActSheetOutput->SetCellValue("A3", "电话：020-86402325");
            $objActSheetOutput->SetCellValue("A4", "供应商编号：0000202778");
            $objActSheetOutput->SetCellValue("A".$write_row, "商品编号");
            $objActSheetOutput->SetCellValue("B".$write_row, "商品名称");
            $objActSheetOutput->SetCellValue("C".$write_row, "数量");
            $objActSheetOutput->SetCellValue("D".$write_row, "进单价");
            $objActSheetOutput->SetCellValue("E".$write_row, "进单金额");

            if (strpos($title, '店别明细') !== false) {//格式A
                $START_ROW = 11;
                $max_row = $curSheetInput->getHighestRow();
                $date = $curSheetInput->getCellByColumnAndRow($C_COL, 4)->getValue();
                $shop = $curSheetInput->getCellByColumnAndRow($C_COL, 8)->getValue();
                $order_sn = str_replace('店别明细', '', $title);//在title拿订单号
                $objActSheetOutput->SetCellValue("A5", "订单号：".$order_sn);
                $objActSheetOutput->SetCellValue("A6", "店名：".$shop);
                $objActSheetOutput->SetCellValue("A7", "送货时间：".$date);
                $all_amount = 0;
                $all_total = 0;

                for ($iRow = $START_ROW; $iRow < $max_row; $iRow++) {
                    $name = $curSheetInput->getCellByColumnAndRow($C_COL+2, $iRow)->getValue();
                    if (empty($name)) continue;
                    $write_row++;
                    $temp = array(
                        'sn'=>$curSheetInput->getCellByColumnAndRow($C_COL, $iRow)->getValue(),
                        'name'=>$name,
                        'price'=>$curSheetInput->getCellByColumnAndRow($C_COL+4, $iRow)->getValue(),
                        'total'=>$curSheetInput->getCellByColumnAndRow($C_COL+5, $iRow)->getValue(),
                        'amount'=>$curSheetInput->getCellByColumnAndRow($C_COL+9, $iRow)->getValue(),
                    );
                    $all_amount += $temp['amount'];
                    $all_total += $temp['total'];
                    $objActSheetOutput->SetCellValue("A".$write_row, $temp['sn']);
                    $objActSheetOutput->SetCellValue("B".$write_row, $temp['name']);
                    $objActSheetOutput->SetCellValue("C".$write_row, $temp['amount']);
                    $objActSheetOutput->SetCellValue("D".$write_row, $temp['price']);
                    $objActSheetOutput->SetCellValue("E".$write_row, $temp['total']);
                }
                $write_row++;
                $objActSheetOutput->SetCellValue("A".$write_row, '合计');
                $objActSheetOutput->SetCellValue("C".$write_row, $all_amount);
                $objActSheetOutput->SetCellValue("E".$write_row, $all_total);
            } else {//格式B
                $START_ROW = 9;
                $max_row = $curSheetInput->getHighestRow();
                $date = $curSheetInput->getCellByColumnAndRow($C_COL+7, 4)->getValue();
                $shop = $curSheetInput->getCellByColumnAndRow($C_COL+1, 5)->getValue();
                $order_sn = $curSheetInput->getCellByColumnAndRow($C_COL+3, 3)->getValue();
                $objActSheetOutput->SetCellValue("A5", "订单号：".$order_sn);
                $objActSheetOutput->SetCellValue("A6", "店名：".$shop);
                $objActSheetOutput->SetCellValue("A7", "送货时间：".$date);
                $all_amount = 0;
                $all_total = 0;

                for ($iRow = $START_ROW; $iRow < $max_row; $iRow++) {
                    $name = $curSheetInput->getCellByColumnAndRow($C_COL+1, $iRow)->getValue();
                    if (empty($name)) continue;
                    $write_row++;
                    $temp = array(
                        'sn'=>$curSheetInput->getCellByColumnAndRow($C_COL, $iRow)->getValue(),
                        'name'=>$name,
                        'price'=>$curSheetInput->getCellByColumnAndRow($C_COL+10, $iRow)->getValue(),
                        'total'=>$curSheetInput->getCellByColumnAndRow($C_COL+11, $iRow)->getValue(),
                        'amount'=>$curSheetInput->getCellByColumnAndRow($C_COL+8, $iRow)->getValue(),
                    );
                    $all_amount += $temp['amount'];
                    $all_total += $temp['total'];
                    $objActSheetOutput->SetCellValue("A".$write_row, $temp['sn']);
                    $objActSheetOutput->SetCellValue("B".$write_row, $temp['name']);
                    $objActSheetOutput->SetCellValue("C".$write_row, $temp['amount']);
                    $objActSheetOutput->SetCellValue("D".$write_row, $temp['price']);
                    $objActSheetOutput->SetCellValue("E".$write_row, $temp['total']);
                }
                $write_row++;
                $objActSheetOutput->SetCellValue("A".$write_row, '合计');
                $objActSheetOutput->SetCellValue("C".$write_row, $all_amount);
                $objActSheetOutput->SetCellValue("E".$write_row, $all_total);
            }

            //调整样式
            $objActSheetOutput->getDefaultStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $objActSheetOutput->getStyle('A2:A7')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
            $objActSheetOutput->getDefaultStyle()->getFont()->setName('simsun')->setSize(12);//宋体,大小12
            $objActSheetOutput->mergeCells('A1:E1');//合并第一行（合并完后可以用A1表示）
            $objActSheetOutput->getStyle('A1')->getFont()->setSize(14);
            //$objActSheetOutput->getStyle('A2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);//左右居中（默认靠左）
            //$objActSheetOutput->getStyle('A3')->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//上下居中（默认靠下）
            $objActSheetOutput->getDefaultRowDimension()->setRowHeight(20);//默认20高度
            $objActSheetOutput->getRowDimension(1)->setRowHeight(25);//第一行25高度
            $objActSheetOutput->getStyle('A8:E'.$write_row)->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
            $objActSheetOutput->getColumnDimension('A')->setWidth(12);//第一列宽12
            $objActSheetOutput->getColumnDimension('B')->setWidth(42);
            $objActSheetOutput->getColumnDimension('C')->setWidth(9);
            $objActSheetOutput->getColumnDimension('D')->setWidth(9);
            $objActSheetOutput->getColumnDimension('E')->setWidth(13);
        }

        $filename = 'dianbie';
        $objWriter = \PHPExcel_IOFactory::createWriter($objOutput, 'Excel5');
        header('Content-Type: application/vnd.ms-excel');
        header("Content-Disposition:attachment; filename=".urlencode($filename).".xls");
        header('Cache-Control: max-age=0');
        $objWriter->save('php://output');
        exit;
        return Yii::$app->response->sendContentAsFile('123','test.txt');
    }

}
