<?php

//create a unique id
$up_id = uniqid();
?>
<?php

$upload_progress_js = <<<UPJS

    jQuery(function($) {
      $('#SubmissionAddForm').submit(function(e) {
		  	show_progress();
      });
 	});

function show_progress(){

		$('.submit').hide();	
		$('#upload_frame').show(); 

		function set() { 
            $('#upload_frame').attr('src','/progress_frame.php?up_id=$up_id'); 
        } 
        setTimeout(set);

}

UPJS;
?>
<!--<form id="UploadAddFile" enctype="multipart/form-data" method="post" action="/upload/addFile">-->
<?php

$up_p_name = ini_get("session.upload_progress.name");
echo $this->Form->create('Upload', array('controller' => 'Upload','action' => 'addFile','type' => 'file', 'default' => false));

$this->Form->unlockField($up_p_name);
echo $this->Form->hidden($up_p_name, array('value' => $up_id, 'name' => $up_p_name, 'id' => 'upload_progress_id', 'secure' => false));
echo $this->Form->input('filename', array('type' => 'file', 'label' => __('File: ')));
echo $this->Form->input(__('Upload File'), array(
                'label' => FALSE,
                'div' => FALSE,
                'type' => 'submit',
                'class' => 'btn bg-darkCyanDark',
                'formnovalidate' => TRUE,
                'id' => 'submit_upload_submitbtn'
                    )
            );
?>

<iframe id="upload_frame" name="upload_frame" frameborder="0" border="0" src="" scrolling="no" scrollbar="no" > </iframe>

<?php

$this->Js->buffer($upload_progress_js);
echo $this->Js->writeBuffer();
?>