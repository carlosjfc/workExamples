<?php
/**
 * Created by PhpStorm.
 * User: carlos.fernandez1
 * Date: 10/13/16
 * Time: 12:21 PM
 */

namespace McKesson;
require_once __DIR__ . '/../../autoload.php';

use McKesson\Utils;
use McKesson\XMLStrings;
use McKesson\ConnectionService;

class McKesson {
    /**
     * the path to the Web service's configuration setting
     */
    const CONFIG_SETTINGS = __DIR__ . '/../config/settings.ini';
    /**
     * the path to the Web service's messages from FIA and Copy department
     */
    const CONFIG_MESSAGES = __DIR__ . '/../config/messages.ini';

    /**
     * @var string the right end point to be connecting depending on the environment
     */
    private $endPoint;
    /**
     * @var string the right password to be knock the end point door depending on the environment
     */
    private $password;

    /**
     * McKesson constructor.
     * @param $endPoint
     * @param $action
     * @param $password
     */
    public function __construct() {
        if (!defined('INITIATIVE_ID')) {
            Utils::loadConsts(McKesson::CONFIG_SETTINGS);
        }
        if (!defined('E000_COPY')) {
            Utils::loadConsts(McKesson::CONFIG_MESSAGES);
        }
        var_dump(self::defineMode());
        $this->endPoint = (self::defineMode() === 'prod') ? MCKESSON_PROD_ENDPOINT : MCKESSON_UAT_ENDPOINT;
        $this->password = (self::defineMode() === 'prod') ? MCKESSON_PROD_PASSWORD : MCKESSON_UAT_PASSWORD;
    }


    /** We have to determinate if we are runing on production|development|stage| other testing server ex: localhost
     * @param bool $force_mds_uat if we want to force testing
     * @return string the enviroment for the server execute the calls against production or uat end points
     */
    private static function defineMode($force_mds_uat = FALSE) {
        if ($force_mds_uat) {
            return MDS_MODE_UAT;
        }
        $server = filter_input(INPUT_SERVER, 'SERVER_NAME', FILTER_UNSAFE_RAW);
        if ($server === PROD_SERVER || $server === 'www.' . PROD_SERVER) {
            return MDS_MODE_PROD;
        }
        return MDS_MODE_UAT;
    }


    /** This method will parse the response from the Web Service and return the response that we will provide to analize success or error on the call
     * @param $response the XML returned by the SOAP Web Service when resquested by the end point and using our class ConnectionService
     * @param $action cardValidation|requestCard
     * @return array array always with the strcuture of sucess => true|false; code => the success|error code that match with McKesson documentation; returnedCardId|error depending on success or error this will contain the right value representing the cardID or the message that is related to the returned error
     */
    private function parseResponse($response, $action) {
        $doc = new DOMDocument();
        $doc->loadXML($response);
        $codeNode = $doc->getElementsByTagName("Code");
        $successResponse = '';
        switch ($action) {
            case 'cardValidation':
                $successResponse = CARD_VALIDATION_RESPONSE;
                break;
            case 'requestCard':
                $successResponse = REGISTRATION_ENROLLMENT_RESPONSE;
                break;
        }

        if ($codeNode->length != 0) {
            $code = $codeNode->item(0)->nodeValue;
            if ($code === $successResponse) {
                $cardIDNode = $doc->getElementsByTagName("CardID");
                //returning a success with the right code and the returned cardID by the server.
                //If is a new card ID request it will be a new one because the original cardID was empty,
                //if we are activating a card this cardID from the server will be the same as the original that we pass by paarmeters, menaing the same that we capture on the forms
                array('success' => TRUE, 'code' => $code, 'returnedCardId' => $cardIDNode->item(0)->nodeValue);
            } else {
                //returning a no success with the error code returned by the server
                return array('success' => FALSE, 'code' => $code, 'error' => constant($code . '_' . INDEX_ERRORS));
            }
        }
        //final return of non success
        return array('success' => FALSE, 'code' => null, 'error' => 'It appears that the server returns an empty XML response');
    }


    /** This one is the only exposed method and the only one that can be called after we instanciate the McKesson class. Encapsules all the rest of the logic to interact with McKesson SOAP WebServives
     * @param $service cardValidation|requestCard
     * @param $data the data array of the fields capture on the forms
     * @return array always with the strcuture of sucess => true|false; code => the success|error code that match with McKesson documentation; returnedCardId|error depending on success or error this will contain the right value representing the cardID or the message that is related to the returned error
     */
    public function callService($service = 'cardValidation', $data) {
        $cardId = (isset($data['cardId'])) ? $data['cardId'] : '';
        if (!empty($cardId) && (strlen($cardId) != 9 || !is_numeric($cardId))) {
            return array('success' => FALSE, 'code' => 'E503', 'error' => E503_COPY);
        }
        $enpointMcKesson = $this->endPoint . "cardValidation?wsdl";
        $actionMcKesson = $this->endPoint . "cardValidation";
        //to the cardValidation action we only need to pass the password of the environment to create the credentials, and the cardId to validate from the FIA page 5
        $XMLString = XMLStrings::validateCard($this->password, $cardId);
        if ($service === 'requestCard') {
            $enpointMcKesson = $this->endPoint . "registration?wsdl";
            $actionMcKesson = $this->endPoint . "registration";
            //to the requestCard action we need quite more fields as per the documentation from the FIA page 7
            //the $data associative array is build from the form captured on the front end
            //so, in the same part of the code where you need to instanciate this class you create this $data array and map the fields with the values from the form
            $Q_FIRSTNAME = (isset($data['firstName'])) ? $data['firstName'] : '';
            $Q_LASTNAME = (isset($data['lastName'])) ? $data['lastName'] : '';
            $Q_ADDRESS1 = (isset($data['address1'])) ? $data['address1'] : '';
            $Q_ADDRESS2 = (isset($data['address2'])) ? $data['address2'] : '';
            $Q_CITY = (isset($data['city'])) ? $data['city'] : '';
            $Q_STATE = (isset($data['state'])) ? $data['state'] : '';
            $Q_ZIP = (isset($data['zip'])) ? $data['zip'] : '';
            $Q_PHONE = (isset($data['phone'])) ? $data['phone'] : '';
            $Q_EMAIL = (isset($data['email'])) ? $data['email'] : '';
            $Q_DOB = (isset($data['dob'])) ? $data['dob'] : '';
            $Q_COUNTRY = (isset($data['country'])) ? $data['country'] : '';
            $Q_GOVT_PAID = (isset($data['govtPaid'])) ? $data['govtPaid'] : '';
            $Q_MED_ACK = (isset($data['medAck'])) ? $data['medAck'] : '';
            $Q_INSURANCE = (isset($data['insurance'])) ? $data['insurance'] : '';
            $Q_RECEIVE_INFO = (isset($data['receiveInfo'])) ? $data['receiveInfo'] : '';
            $Q_MARKETING_OPT_IN = (isset($data['marketingOptIn'])) ? $data['marketingOptIn'] : '';
            $Q_LANGUAGE = (isset($data['language'])) ? $data['language'] : '';
            $Q_HOW_LONG = (isset($data['howLong'])) ? $data['howLong'] : '';
            $Q_INSUR_RED = (isset($data['insurRed'])) ? $data['insurRed'] : '';
            $Q_SOURCE = (isset($data['source'])) ? $data['source'] : '';
            //we will pass the password for the environment to create the credentials together with the fields as per the documentation from the FIA page 7
            $XMLString = XMLStrings::enrollmentenrollment($this->password, $cardId, $Q_FIRSTNAME, $Q_LASTNAME, $Q_ADDRESS1, $Q_ADDRESS2, $Q_CITY, $Q_STATE, $Q_ZIP, $Q_PHONE, $Q_EMAIL, $Q_DOB, $Q_COUNTRY, $Q_GOVT_PAID, $Q_MED_ACK, $Q_INSURANCE, $Q_RECEIVE_INFO, $Q_MARKETING_OPT_IN, $Q_LANGUAGE, $Q_HOW_LONG, $Q_INSUR_RED, $Q_SOURCE);
        }
        $connectionService = new ConnectionService($enpointMcKesson, null, null, $actionMcKesson, $XMLString);
        $response = $connectionService->request();
        return $this->parseResponse($response, 'validateCard');
    }

}