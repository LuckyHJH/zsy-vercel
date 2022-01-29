
<?php
/* @var $this yii\web\View */

$this->title = 'zsy';

use yii\widgets\ActiveForm;
?>

<?php $form = ActiveForm::begin(['options' => ['enctype' => 'multipart/form-data', 'style' => 'margin-bottom: 80px;']]) ?>

退货单

<?= $form->field($model, 'file')->fileInput() ?>

<input type="hidden" name="type" value="tuihuo">

<button>Submit</button>

<?php ActiveForm::end() ?>


<?php $form = ActiveForm::begin(['options' => ['enctype' => 'multipart/form-data', 'style' => 'margin-bottom: 80px;']]) ?>

<!--装箱单

<?= $form->field($model, 'file')->fileInput() ?>

<input type="hidden" name="type" value="zhuangxiang">

<button>Submit</button>

<?php ActiveForm::end() ?>


<?php $form = ActiveForm::begin(['options' => ['enctype' => 'multipart/form-data', 'style' => 'margin-bottom: 80px;']]) ?>
-->

店别单

<?= $form->field($model, 'file')->fileInput() ?>

<input type="hidden" name="type" value="dianbie">

<button>Submit</button>

<?php ActiveForm::end() ?>