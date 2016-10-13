<?php
/**
 * Created by PhpStorm.
 * User: carlos.fernandez1
 * Date: 10/13/16
 * Time: 12:54 PM
 */
require_once 'autoload.php';
use McKesson\McKesson;

$mk = new McKesson();

//gets here the fileds from the form and create the $data array; for example:
$data = array();
$data['cardId'] = filter_input(INPUT_POST, 'form_card_id_field_name', FILTER_UNSAFE_RAW);
$data['firstName'] = filter_input(INPUT_POST, 'form_first_name_field_name', FILTER_UNSAFE_RAW);
//ETC.....
//this one will return the array resulting on the requested Web services as an JSON array to the AJAX call
return json_encode($mk->callService('requestCard', $data));