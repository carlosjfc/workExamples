<?php
if (!empty($list_files)) {
    ?>
    <div>
        <table class="fuploads table bordered">
            <thead >
                <tr>
                    <th class="text-center small">Type</th>
                    <th class="text-left">Name</th>
                    <th class="text-left" id="header_file_actions">Actions</th>
                </tr>
            </thead>
            <tbody>
                <?php
                $before = $this->Js->get('#ajax_transiction_w8_img')->effect('fadeIn', array('buffer' => false));
                $success = $this->Js->get('#ajax_transiction_w8_img')->effect('fadeOut', array('buffer' => false));
                $referer = $this->request->referer();
                $slaches = substr_count($referer, '/');
                $slaches -= 4;
                $pre = '';
                for ($i = 0; $i < $slaches; $i++)
                    $pre .= '../';
                foreach ($list_files as $key => $files) {
                    foreach ($files as $key => $file) {
                        $this->Js->get('#deleteUploadedFile_' . $file['id'])->event(
                                'click', $this->Js->request(
                                        array('action' => 'deleteUploadedFile', $file['id']), array(
                                    'success' => '$.post(\'' . $pre . 'uploads/listFiles\', function(data) {
                                                                                                        $(\'#files_table\').html(data);
                                                                                                        document.getElementById(\'ajax_transiction_w8_img\').style.visibility = \'hidden\';
                                                                                                    })',
                                    'async' => true,
                                    'before' => $before,
                                    'dataExpression' => true,
                                    'method' => 'POST'
                                        )
                                )
                        );
                        $icon_type = 'icon-file';
                        $icon_color = 'fg-darkBlue';
                        $fname = $file['fname'];
                        $pieces = explode(".", $fname);
                        $extension = $pieces[count($pieces) - 1];
                        if ($extension == 'xls' || $extension == 'xlsx' || $extension == 'xlsb' || $extension == 'xlsm' || $extension == 'xlt' || $extension == 'xltx' || $extension == 'xlw') {
                            $icon_type = 'icon-file-excel';
                            $icon_color = 'fg-green';
                        } elseif ($extension == 'doc' || $extension == 'docx' || $extension == 'dot' || $extension == 'dotx') {
                            $icon_type = 'icon-file-word';
                            $icon_color = 'fg-darkBlue';
                        } elseif ($extension == 'pot' || $extension == 'pps' || $extension == 'ppt' || $extension == 'pptx') {
                            $icon_type = 'icon-file-powerpoint';
                            $icon_color = 'fg-orangeDark';
                        } elseif ($extension == 'zip') {
                            $icon_type = 'icon-file-zip';
                            $icon_color = 'fg-yellow';
                        } elseif ($extension == 'pdf') {
                            $icon_type = ' icon-file-pdf';
                            $icon_color = 'fg-darkMagenta';
                        }
                        echo '<tr>';
                        echo "<td class ='center text-center small'><i class = '$icon_type $icon_color'></i></td>";
                        echo '<td class =" text-left">' . $this->Html->link($file['fname'], '../' . $file['filename'], array('target' => '_blank')) . '</td>
                                <td id="cell_file_actions_' . $file['id'] . '" class ="cell_file_action text-left">';
                        echo $this->Html->tag('i', '', array(
                            'class' => 'icon-remove fg-red cursor-hand',
                            'title' => 'Delete File',
                            'alt' => 'Delete File',
                            'id' => 'deleteUploadedFile_' . $file['id']
                        ));
                        echo '</td></tr>';
                    }
                }
                ?>
            </tbody>
        </table>
    </div>
    <?php
}
?>
<?php
$scripts = <<<UPJS
    $(document).ready(function(){
        $('.fuploads').css({'min-width':'100%','background-color':'transparent'});
    });

UPJS;
?>
<?php $this->Js->buffer($scripts); ?>
<?php echo $this->Js->writeBuffer(); ?>