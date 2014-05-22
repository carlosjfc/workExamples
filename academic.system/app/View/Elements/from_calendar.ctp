<?php
App::uses('CakeSession', 'Model/Datasource');
?>

<!--contenido-->
<div class="grid">
    <div class="row">
        <div class="span4 required" id="calendarFromDiv">
            <?php
            /*
             * Incluir el calendario y cuando cambie el valor llamar a getHeadCountsRoster($date)
             */
            $btn_date = $this->Html->tag('button', NULL, array('class' => 'btn-date', 'type' => 'button'));
            $btn_clear = $this->Html->tag('button', NULL, array('class' => 'btn-clear', 'tabindex' => '-1', 'type' => 'button'));
            //$options_date = array('data-param-lang' => 'en', 'data-role' => 'datepicker', 'data-param-year-buttons' => 1);
            $options_date = array('data-role' => "calendar");


            //$date_div = $this->Html->div('input-control text datepicker fromdate required', $date . $btn_date, $options_date);
            //$date_div = $this->Html->div('calendar small', NULL, $options_date);
            $date_div = $this->Html->div('calendar calendarfrom small', NULL, NULL);
            echo $date_div;
            $date_from = $this->Form->hidden('date_from', array(
                'type' => 'text',
                'id' => 'date_from',
                'label' => FALSE,
                'div' => FALSE,
                'placeholder' => 'Date',
                'autofocus' => TRUE,
                'required' => TRUE,
                'enabled' => FALSE
                    )
            );
            ?>
        </div>
    </div>
    <?php
    echo $date_from;
    ?>

<div class="span4" id="action_btn">
<?php
echo $this->Form->button('Pull Report', array(
    'label' => FALSE,
    'div' => FALSE,
    'type' => 'button',
    'class' => 'btn  default',
    'formnovalidate' => FALSE,
    'id' => 'search_btn',
    'style' => 'position:relative;top:33px;'
        )
);
?>
</div>
</div>
</div>
<?php
$calendar_setup = <<<UCL
$(document).ready(function(){
    $(function(){
                            var calfrom = $(".calendarfrom").calendar({
                                    format: 'yyyy-mm-dd',
                                    multiSelect: false,
                                    /*getDates: function(data){
                                        var r = "", out = $("#calendar-output").html("");
                                        $.each(data, function(i, d){
                                        r += d + "<br />";
                                        });
                                        out.html(r);
                                    },*/
                                    click: function(d){
                                        var out = $("#date_from").val("");
                                        out.val(d);
                                        //$('#HeadCountGetHeadCountsRosterForm').submit();
                                    }
                            });
                        });
        $('.calendarfrom').before('<label for="valid_from">Date</label>');
        $('#date_from').appendTo('#calendarFromDiv');
        $('.metro form').css({'margin':'0 0 0'});
        $('.metro .grid .row').css({'margin-top':'0px'});
        $('.metro .calendar.small .day-of-week, .metro .calendar.small a').addClass('fg-black');
        $('#search_btn').click(function(){
            var date_from = $("#date_from").val();
            var errors = false;
            if(date_from == ''){
                $.Dialog({
                            shadow: true,
                            overlay: false,
                            icon: '<span class="icon-cancel-2 fg-darkRed"></span>',
                            title: 'Errors',
                            width: 500,
                            padding: 0,
                            content: 'The Date have to be selected'
                        });
             errors = true;
            }
            if(!errors){
                $('#ajax_transiction_w8_img').fadeIn();
                var calendar_ajax_function = '$calendar_ajax_function';
                $.post(calendar_ajax_function+'/'+date_from, function(data) {
                                                    $('#report_area').html(data);
                                                    $('#ajax_transiction_w8_img').fadeOut();
                                         });
            }
            $('.btn-close').click(function(){ $('#ajax_transiction_w8_img').fadeOut(); });
        });
});
UCL;
?>
<?php
$this->Js->buffer($calendar_setup);
echo $this->Js->writeBuffer();
?>