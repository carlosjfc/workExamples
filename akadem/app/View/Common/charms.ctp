<?php $this->start('charms'); ?>
<?php
$charms = $this->requestAction('charms/index/');
$charm_bg_color = $charms['charm_bg_color'];
$charm_fg_color = $charms['charm_fg_color'];
?>
<div class="charms <?php if (isset($charm_bg_color))
    echo $charm_bg_color; echo ' ';
if (isset($charm_fg_color))
    echo $charm_fg_color;
?> transparent">
    <div>
<?php echo $this->fetch('charms-content'); ?>
    </div>
</div>
<?php $this->end(); ?>
<?php echo $this->fetch('page-content'); ?>
<?php $this->start('page_script') ?>
<script>
    $(document).ready(function() {
        $('.charms').hover(function() {
            $(this).addClass('open');
            $(this).removeClass('transparent');
        }, function() {
            $(this).removeClass('open');
            $(this).addClass('transparent');
        });
    });
</script>
<?php $this->end(); ?>
<?php echo $this->fetch('script-content'); ?>