<?php

/*
 * This controller will be handling all the requests to work with uploading files.
 * Basicly all the work will be doing by AJAX calls
 * @author CARLOS FERNANDEZ
 * @email carlosjfc@gmail.com
 * @company Drake College of Business
 */
?>
<?php

class UploadsController extends AppController {

    public $name = 'Uploads';
    public $uses = array('Uploads');

    /**
     * Add an uploaded file to the file system and the database
     */
    public function addFile() {
        $this->layout = 'ajax';
        if ($this->request->is('post') || $this->request->is('ajax')) {
            $this->Uploads->create();
            $this->Uploads->set('dd_id', $this->UserAuth->getUserDDId());
            if ($this->Uploads->save($this->request->data)) {
                $this->Session->setFlash(__('The uploaded file has been saved'));
                $result = 1;
                $this->set(compact('result'));
            } else {
                $this->Session->setFlash(__('The uploaded file couldn\'t been saved'));
            }
        }
    }

    /**
     * Get the list of file uploaded for the current user's session to the active form 
     */
    public function listFiles() {
        $this->layout = 'ajax';
        $form_random_id = $this->Session->read('form_random_id');
        $list_files = $this->Uploads->getUploadedFiles($form_random_id);
        $this->set(compact('list_files'));
    }

    /**
     * Get the list of the uploaded files for an specif model/model_id pair
     * @param string $model the model name used to create and upload the Upload Form view
     * @param int $model_id the id generated to the specif $model to bind with created form on the view
     */
    public function listModelFiles($model, $model_id) {
        $this->layout = 'ajax';
        $list_files = $this->Uploads->getUploadedFilesFromId($model_id, $model);
        $this->set(compact('list_files'));
        $this->render('list_files_no_actions');
    }

    /**
     * Get the list of the uploaded files for an specif model/model_id pair and update the new form_id;
     * that form_id will be the responsable for the new edit/actions that will take place with all the fetched files from the pair $model/$model_id
     * @param string $model the model name used to create and upload the Upload Form view
     * @param int $model_id the id generated to the specif $model to bind with created form on the view
     */
    public function listEditFiles($model, $model_id) {
        $this->layout = 'ajax';
        $form_random_id = $this->Session->read('form_random_id');
        $this->Uploads->editFormUploadId($form_random_id, $model, $model_id);
        $list_files = $this->Uploads->getUploadedFiles($form_random_id, $model);
        $this->set(compact('list_files'));
        $this->render('listFiles');
    }

    /**
     * 
     * @param int $file_id the file id to delete
     */
    public function deleteUploadedFile($file_id) {
        $this->layout = 'ajax';
        $this->autoRender = FALSE;
        $this->Uploads->deleteUploadedFile($file_id);
    }

    public function index() {
        $this->layout = 'w8_frame_page';
    }

    /**
     * Check the upload progress for a specific file
     * @return string|float the upload progress of the file
     */
    public function uploadProgress() {
        if ($this->request->is('ajax')) {
            $this->autoRender = FALSE;
            $unique_upload_id = $_POST['up_id'];
            $key = ini_get("session.upload_progress.prefix") . $unique_upload_id;
            if (isset($_SESSION[$key])) {
                $upload_progress = $_SESSION[$key];
                $progress = round(($upload_progress['bytes_processed'] / $upload_progress['content_length']) * 100, 0);
                return $progress;
            } else {
                return "100000";
            }
        }
    }

}

?>