<?php
App::uses('CakeSession', 'Model/Datasource');
?>
<?php
echo $this->Html->script('metro/metro/metro-input-control');
?>
<style>
    #f1_upload_process{
        z-index:100;
        position:absolute;
        visibility:hidden;
        text-align:center;
        width:400px;
        margin:0px;
        padding:0px;
        /*background-color:#fff;*/
        border:1px solid #ccc;
    }
    div.custom_file_upload {
        width: 230px;
        height: 20px;
        /*margin: 40px auto;*/
        margin: -15px 0 20px;
    }

    input.file {
        width: 150px;
        height: 20px;
        border: 1px solid #BBB;
        border-right: 0;
        color: #888;
        padding: 5px;

        -webkit-border-top-left-radius: 5px;
        -webkit-border-bottom-left-radius: 5px;
        -moz-border-radius-topleft: 5px;
        -moz-border-radius-bottomleft: 5px;
        border-top-left-radius: 5px;
        border-bottom-left-radius: 5px;

        outline: none;
    }

    div.file_upload {
        width: 80px;
        height: 24px;
        background: #7abcff;
        background: -moz-linear-gradient(top,  #7abcff 0%, #60abf8 44%, #4096ee 100%);
        background: -webkit-gradient(linear, left top, left bottom, color-stop(0%,#7abcff), color-stop(44%,#60abf8), color-stop(100%,#4096ee));
        background: -webkit-linear-gradient(top,  #7abcff 0%,#60abf8 44%,#4096ee 100%);
        background: -o-linear-gradient(top,  #7abcff 0%,#60abf8 44%,#4096ee 100%);
        background: -ms-linear-gradient(top,  #7abcff 0%,#60abf8 44%,#4096ee 100%);
        background: linear-gradient(top,  #7abcff 0%,#60abf8 44%,#4096ee 100%);
        filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#7abcff', endColorstr='#4096ee',GradientType=0 );

        display: inline;
        position: absolute;
        overflow: hidden;
        cursor: pointer;

        -webkit-border-top-right-radius: 5px;
        -webkit-border-bottom-right-radius: 5px;
        -moz-border-radius-topright: 5px;
        -moz-border-radius-bottomright: 5px;
        border-top-right-radius: 5px;
        border-bottom-right-radius: 5px;


        font-weight: bold;
        color: #FFF;
        text-align: center;
        padding-top: 8px;
    }
    div.file_upload:before {
        content: ' ';
        position: absolute;
        left: 0; right: 0;
        text-align: center;
        cursor: pointer;
    }

    div.file_upload input {
        position: relative;
        height: 30px;
        width: 250px;
        display: inline;
        cursor: pointer;
        opacity: 0;
    }
</style>
<?php
//create a unique id
$up_id = uniqid();
?>

<p id="f1_upload_process" class="<?php echo $this->Session->read('bg_color'); ?>" style =" border: none; position: relative; top: -18px; ">Uploading File...<br/>
    <?php
//echo $this->Html->image('preloaders/preloader_bar_dots_blueLight.gif');
    ?>
    <?php
    $bg_parts = explode('-', $this->Session->read('bg_color'));
    $preimg = 'preloaders/preloader_bar_dots_';
    $img_color = $bg_parts[count($bg_parts) - 1];
    $postimg = '.gif';
    $image = $preimg . $img_color . $postimg;
    echo $this->Html->image($image);
    ?>
</p>
<p id="result"></p>
<!--<img src="loader.gif" /></p>-->
<?php
$up_p_name = ini_get("session.upload_progress.name");
echo $this->Form->create('Uploads', array('controller' => 'Uploads', 'action' => 'addFile', 'target' => 'upload_target', 'type' => 'file', 'onsubmit' => 'startUpload()'/* , 'default' => false */));

$this->Form->unlockField($up_p_name);
CakeSession::write('form_random_id', String::uuid());
echo $this->Form->hidden($up_p_name, array('value' => $up_id, 'name' => $up_p_name, 'id' => 'upload_progress_id', 'secure' => false));
echo $this->Form->hidden('form_random_id', array('value' => $this->Session->read('form_random_id'), 'id' => 'form_random_id'));
$model = Inflector::classify($this->params['controller']);
/*if(isset(CakeSession::read('upload_model_write')) && CakeSession::read('upload_model_write') != NULL && CakeSession::read('upload_model_write') != '')
      $model = CakeSession::read('upload_model_write');*/
echo $this->Form->hidden('model', array('value' => $model, 'id' => 'model'));
//$action = $this->action;
//debug($action);
//echo $this->Form->hidden('action', array('value' => $action, 'id' => 'action'));
?>

<div data-role="input-control" class="input-control file">
    <input type="file" style="z-index: 0;" tabindex="-1">
    <button class="btn-file" type="button"></button> 
</div>
<!--<i class="icon-upload fg-darkBlueLight" style="position: absolute;z-index: 1000; font-size:26px; cursor: pointer;"></i>-->
<!-- <i class="icon-upload fg-darkBlueLight" style="position: relative;z-index: 1000; left: 185px; top: -40px; font-size:26px; cursor: pointer;"></i> -->
<?php
//echo $this->Form->end(__('Upload File'));
echo '<p></p>';
echo $this->Form->input(__('Upload File'), array(
    'label' => FALSE,
    'div' => FALSE,
    'type' => 'submit',
    'class' => 'btn  default',
    'formnovalidate' => TRUE,
    'id' => 'submit_upload_submitbtn'
        )
);
?>
<iframe id="upload_target" name="upload_target" src="#" style="width:0;height:0;border:0px solid #fff;"></iframe>
<div id="files_table"></div>
<?php
$referer = $this->request->referer();
$slaches = substr_count($referer, '/');
$slaches -= 4;
$pre = '';
for($i = 0; $i < $slaches; $i++) $pre .= '../';
?>
<?php
$upload_custom_last = <<<UCL
$(document).ready(function(){
    $('input[type="file"]').attr('name','data[Uploads][filename]');
});
function startUpload(){
    document.getElementById('f1_upload_process').style.visibility = 'visible';
    return true;
}        
function stopUpload(success){
      var result = '';
      if (success == 1 || success == '1'){
         /*document.getElementById('result').innerHTML =
           '<span class="msg">The file was uploaded successfully!<\/span><br/><br/>';*/
           /*Trigger the Ajax to load the files table*/
           var pre = '$pre';
           $.post(pre+'uploads/listFiles', function(data) {
                                                    $('#files_table').html(data);
                                                    document.getElementById('f1_upload_process').style.visibility = 'hidden';
           });
      }
      else {
         document.getElementById('result').innerHTML = 
           '<span class="emsg">There was an error during file upload!<\/span><br/><br/>';
           document.getElementById('f1_upload_process').style.visibility = 'hidden';
      }
      return true;   
}

UCL;
?>
<?php
$this->Js->buffer($upload_custom_last);
echo $this->Js->writeBuffer();
?>