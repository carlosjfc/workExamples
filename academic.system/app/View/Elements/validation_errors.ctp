<?php
if (!empty($errors)) {
    ?>
    <div class="message-dialog bg-red fg-white" style="position:relative;">
        <p class="fg-white"><?php echo $error_msg; ?></p>
    </div>
    <?php
}
?>
<?php
$this->Js->buffer('$(\'.error_msg\').remove();');
if (!empty($errors)) {
    $inputs = array_keys($errors);
    //debug($inputs);
    foreach ($inputs as $key => $input) {
        $this->Js->buffer('$(\'<p class="error_msg fg-darkMagenta">' . $errors[$input][0] . '</p>\').insertAfter(\'#' . $input . '\');');
    }
}
?>
<?php
echo $this->Js->writeBuffer();
?>