<?php

/*
 * The Upload Model will handle all the data validation and CRUD operations with the DB
 * and file system operations to upload files to the server
 * @author CARLOS FERNANDEZ
 * @email carlosjfc@gmail.com
 * @company Drake College of Business
 */
?>
<?php

App::uses('AttachmentBehavior', 'Uploader.Model/Behavior');

class Uploads extends AppModel {

    public $name = 'Uploads';
    public $useTable = 'DCB_Uploads';
    public $validate = array(
        'filename' => array(
            'uploadError' => array(
                'rule' => 'uploadError',
                'message' => 'Something went wrong with the file upload',
                'required' => FALSE,
                'allowEmpty' => TRUE,
            ),
            // custom callback to deal with the file upload
            'processUpload' => array(
                'rule' => 'processUpload',
                'message' => 'Something went wrong processing your file',
                'required' => FALSE,
                'allowEmpty' => TRUE,
                'last' => TRUE,
            )
        )
    );

    /**
     * Upload Directory relative to WWW_ROOT
     * @param string
     */
    public $uploadDir = 'files/uploads';

    /**
     * Process the Upload
     * @param array $check
     * @return bool
     */
    public function processUpload($check = array()) {
        // deal with uploaded file
        if (!empty($check['filename']['tmp_name'])) {

            // check file is uploaded
            if (!is_uploaded_file($check['filename']['tmp_name'])) {
                return FALSE;
            }

            // build full filename
            $dd_id = $this->data[$this->alias]['dd_id'];
            $time = date('Y_m_d_h_i_s');
            $filename = WWW_ROOT . $this->uploadDir . DS . $dd_id . '_' . $time . '_' . Inflector::slug(pathinfo($check['filename']['name'], PATHINFO_FILENAME)) . '.' . pathinfo($check['filename']['name'], PATHINFO_EXTENSION);
            $this->data[$this->alias]['fname'] = Inflector::slug(pathinfo($check['filename']['name'], PATHINFO_FILENAME)) . '.' . pathinfo($check['filename']['name'], PATHINFO_EXTENSION);
            $this->data[$this->alias]['uploaded'] = date('Y-m-d h:i:s');

            // @todo check for duplicate filename
            // try moving file
            if (!move_uploaded_file($check['filename']['tmp_name'], $filename)) {
                return FALSE;

                // file successfully uploaded
            } else {
                // save the file path relative from WWW_ROOT e.g. uploads/example_filename.jpg
                $this->data[$this->alias]['filepath'] = str_replace(DS, "/", str_replace(WWW_ROOT, "", $filename));
            }
        }

        return TRUE;
    }

    /**
     * Before Save Callback
     * @param array $options
     * @return bool
     */
    public function beforeSave($options = array()) {
        // a file has been uploaded so grab the filepath
        if (!empty($this->data[$this->alias]['filepath'])) {
            $this->data[$this->alias]['filename'] = $this->data[$this->alias]['filepath'];
        }

        return TRUE;
    }

    /**
     * Before Validation
     * @param array $options
     * @return bool
     */
    public function beforeValidate($options = array()) {
        parent::beforeSave($options);

        // ignore empty file - causes issues with form validation when file is empty and optional
        if (!empty($this->data[$this->alias]['filename']['error']) && $this->data[$this->alias]['filename']['error'] == 4 && $this->data[$this->alias]['filename']['size'] == 0) {
            unset($this->data[$this->alias]['filename']);
        }

        return TRUE;
    }

    /**
     * Get the list of uploaded files for a specific form
     * @param string $form_random_id the generated form id
     * @return array 
     */
    public function getUploadedFiles($form_random_id) {
        $fields = array('id', 'filename', 'fname', 'model', 'form_random_id');
        $options = array(
            'conditions' => array(
                'AND' => array(
                    'form_random_id' => $form_random_id
                )
            ),
            'fields' => $fields,
            'order' => 'fname'
        );
        return $this->find('all', $options);
    }

    /**
     * Get the list of the uploaded files for an specif model/model_id pair
     * @param int $model_id the id generated to the specif $model to bind with created form on the view
     * @param string $model the model name used to create and upload the Upload Form view
     * @return array
     */
    public function getUploadedFilesFromId($model_id, $model) {
        $fields = array('id', 'filename', 'fname', 'model', 'form_random_id');
        $options = array(
            'conditions' => array(
                'AND' => array(
                    'model_id' => $model_id,
                    'model' => $model
                )
            ),
            'fields' => $fields,
            'order' => 'fname'
        );
        return $this->find('all', $options);
    }

    /**
     * Edit into the database all the values for the parametrs.
     * It's handy when we have to edit the files for a pair $model/$model_id using a new form
     * @param string $form_random_id the generated form id
     * @param string $model the model name used to create and upload the Upload Form view
     * @param int $model_id the id generated to the specif $model to bind with created form on the view
     */
    public function editFormUploadId($form_random_id, $model, $model_id) {
        $sql = "UPDATE [DCB_Uploads] SET [form_random_id] = '$form_random_id'  WHERE [model] = '$model' AND [model_id] = '$model_id'";
        $this->query($sql);
    }

    /**
     * Delete an uploaded file from DB and from filesystem
     * @param int $file_id the file id to delete
     */
    public function deleteUploadedFile($file_id) {
        $fields = array('id', 'filename', 'fname', 'model', 'form_random_id');
        $options = array(
            'conditions' => array(
                'AND' => array(
                    'id' => $file_id
                )
            ),
            'fields' => $fields,
            'order' => 'fname'
        );
        $result = $this->find('first', $options);
        $filename = $result['Uploads']['filename'];
        $path = str_replace(DS, "/", str_replace(WWW_ROOT, "", $filename));
        $file = new File($path);
        $file->delete();
        $this->delete($file_id);
    }

}

?>
