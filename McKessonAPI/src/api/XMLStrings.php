<?php
/**
 * Created by PhpStorm.
 * User: carlos.fernandez1
 * Date: 10/13/16
 * Time: 1:44 PM
 */

namespace McKesson;


class XMLStrings {

    /**
     * @param $password the password in correspondence with the environment
     * @return string the credentials that will be Injected on the request's XML calls
     */
    private static function getCredentials($password) {
        return '<ts:Credentials>
                    <ts:InitiativeID>' . INITIATIVE_ID . '</ts:InitiativeID>
                    <ts:Username>' . USERNAME . '</ts:Username>
                    <ts:Password>' . base64_encode($password) . '</ts:Password>
                </ts:Credentials>';
    }

    /**
     * @param $password
     * @param $cardId
     * @return string the XML to pass on the SOAP Web Service call to validate a card
     */
    public static function validateCard($password, $cardId) {
        return '<?xml version="1.0" encoding="UTF-8"?>
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:web="http://ws-loyaltyscript.mckesson.com/">
            <soapenv:Header/>
            <soapenv:Body>
                <ts:CardValidationRequest xmlns:ts="http://ws-loyaltyscript.mckesson.com/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://ws-loyaltyscript.mckesson.com/ CardValidationSchema.xsd">
                    <ts:ReqType>ID-VALIDATION</ts:ReqType>
                    ' . self::getCredentials($password) . '
                    <ts:CardID>' . $cardId . '</ts:CardID>
                    <ts:AlternateID QuestionID="Q_CAM_ID" ts:Datatype="string">
                        <ts:Value>' . CAM_ID . '</ts:Value>
                    </ts:AlternateID>
                </ts:CardValidationRequest>
            </soapenv:Body>
        </soapenv:Envelope>';
    }

    /**
     * @param $password
     * @param string $Q_CARD_ID
     * @param string $Q_FIRSTNAME
     * @param string $Q_LASTNAME
     * @param string $Q_ADDRESS1
     * @param string $Q_ADDRESS2
     * @param string $Q_CITY
     * @param string $Q_STATE
     * @param string $Q_ZIP
     * @param string $Q_PHONE
     * @param string $Q_EMAIL
     * @param string $Q_DOB
     * @param string $Q_COUNTRY
     * @param string $Q_GOVT_PAID
     * @param string $Q_MED_ACK
     * @param string $Q_INSURANCE
     * @param string $Q_RECEIVE_INFO
     * @param string $Q_MARKETING_OPT_IN
     * @param string $Q_LANGUAGE
     * @param string $Q_HOW_LONG
     * @param string $Q_INSUR_RED
     * @param string $Q_SOURCE
     * @return string returns the XML to pass on the SOAP Web Service call to request a  new card or activate a card
     */
    public static function enrollment($password, $Q_CARD_ID = '', $Q_FIRSTNAME = '', $Q_LASTNAME = '', $Q_ADDRESS1 = '', $Q_ADDRESS2 = '', $Q_CITY = '', $Q_STATE = '', $Q_ZIP = '', $Q_PHONE = '', $Q_EMAIL = '', $Q_DOB = '07/04/1976', $Q_COUNTRY = '', $Q_GOVT_PAID = 'N', $Q_MED_ACK = 'Y', $Q_INSURANCE = 'N', $Q_RECEIVE_INFO = 'N', $Q_MARKETING_OPT_IN = '', $Q_LANGUAGE = '', $Q_HOW_LONG = '', $Q_INSUR_RED = '', $Q_SOURCE = '') {

        //Alternatives and optional fields AS per page 9 on FIA documentation
        $cardIDTag = '';
        if ($Q_CARD_ID != '') {
            $cardIDTag = '<ts:Answer QuestionID="Q_CARD_ID" ts:Datatype="CardIDType">
              			<ts:Value>' . $Q_CARD_ID . '</ts:Value>
			   </ts:Answer>';
        }
        $add2Tag = '';
        if ($Q_ADDRESS2 !== '') {
            $add2Tag = '<ts:Answer QuestionID="Q_ADDRESS2" ts:Datatype="AddressLineType">
              <ts:Value>' . $Q_ADDRESS2 . '</ts:Value>
            </ts:Answer>';
        }

        $emailTag = '';
        if ($Q_EMAIL != '') {
            $emailTag = '<ts:Answer QuestionID="Q_EMAIL" ts:Datatype="EmailType">
                <ts:Value>' . $Q_EMAIL . '</ts:Value>
            </ts:Answer>';
        }
        $marketingTag = '';
        if ($Q_MARKETING_OPT_IN != '') {
            $marketingTag = '<ts:Answer QuestionID="Q_MARKETING_OPT_IN" ts:Datatype="YESNOType">
                <ts:Value>' . $Q_MARKETING_OPT_IN . '</ts:Value>
            </ts:Answer>';
        }
        $languageTag = '';
        if ($Q_LANGUAGE != '') {
            $languageTag = '<ts:Answer QuestionID="Q_LANGUAGE" ts:Datatype="string">
                <ts:Value>' . $Q_LANGUAGE . '</ts:Value>
            </ts:Answer>';
        }
        $howlongTag = '';
        if ($Q_HOW_LONG != '') {
            $howlongTag = '<ts:Answer QuestionID="Q_HOW_LONG" ts:Datatype="string">
                <ts:Value>' . $Q_HOW_LONG . '</ts:Value>
            </ts:Answer>';
        }
        $insur_redTag = '';
        if ($Q_INSUR_RED != '') {
            $insur_redTag = '<ts:Answer QuestionID="Q_INSUR_RED" ts:Datatype="YESNOType">
                <ts:Value>' . $Q_INSUR_RED . '</ts:Value>
            </ts:Answer>';
        }
        $sourceTag = '';
        if ($Q_SOURCE != '') {
            $sourceTag = '<ts:Answer QuestionID="Q_SOURCE" ts:Datatype="string">
                <ts:Value>' . $Q_SOURCE . '</ts:Value>
            </ts:Answer>';
        }

        //Hardcoded fileds AS per page 9 on FIA documentation
        $Q_WHICH_MED = '2';
        $whichmedTag = '<ts:Answer QuestionID="Q_WHICH_MED" ts:Datatype="string">
                <ts:Value>' . $Q_WHICH_MED . '</ts:Value>
            </ts:Answer>';

        return '<?xml version="1.0" encoding="UTF-8"?>
      <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:web="http://ws-loyaltyscript.mckesson.com/">
        <soapenv:Header/>
        <soapenv:Body>
          <ts:RegistrationRequest xmlns:ts="http://ws-loyaltyscript.mckesson.com/"
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://wsloyaltyscript.mckesson.com/ RegistrationSchema.xsd">
            <ts:ReqType>ENROLLMENT</ts:ReqType>
            ' . self::getCredentials($password) . '
            <ts:Answer QuestionID="Q_CAM_ID" ts:Datatype="string">
                <ts:Value>' . CAM_ID . '</ts:Value>
            </ts:Answer>'
        . $cardIDTag .
        '<ts:Answer QuestionID="Q_FIRSTNAME" ts:Datatype="NameType">
                <ts:Value>' . $Q_FIRSTNAME . '</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_LASTNAME" ts:Datatype="NameType">
                <ts:Value>' . $Q_LASTNAME . '</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_ADDRESS1" ts:Datatype="AddressLineType">
                <ts:Value>' . $Q_ADDRESS1 . '</ts:Value>
            </ts:Answer>'
        . $add2Tag .
        '<ts:Answer QuestionID="Q_CITY" ts:Datatype="CityType">
                <ts:Value>' . $Q_CITY . '</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_STATE" ts:Datatype="StateCodeType">
                <ts:Value>' . $Q_STATE . '</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_ZIP" ts:Datatype="ZipType">
                <ts:Value>' . $Q_ZIP . '</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_PHONE" ts:Datatype="PhoneType">
                <ts:Value>' . $Q_PHONE . '</ts:Value>
            </ts:Answer>'
        . $emailTag .
        '<ts:Answer QuestionID="Q_DOB" ts:Datatype="DateType">
                <ts:Value>' . $Q_DOB . '</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_PIN" ts:Datatype="PinType">
                <ts:Value>' . substr($Q_PHONE, -4) . '</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_COUNTRY" ts:Datatype="YESNOType">
                <ts:Value>Y</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_GOVT_PAID" ts:Datatype="YESNOType">
                <ts:Value>' . $Q_GOVT_PAID . '</ts:Value>
            </ts:Answer>
           <ts:Answer QuestionID="Q_MED_ACK" ts:Datatype="YESNOType">
                <ts:Value>' . $Q_MED_ACK . '</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_INSURANCE" ts:Datatype="YESNOType">
                <ts:Value>' . $Q_INSURANCE . '</ts:Value>
            </ts:Answer>
            <ts:Answer QuestionID="Q_RECEIVE_INFO" ts:Datatype="YESNOType">
                <ts:Value>' . $Q_RECEIVE_INFO . '</ts:Value>
            </ts:Answer>'
        . $marketingTag . $languageTag . $howlongTag . $insur_redTag . $whichmedTag . $sourceTag .
        '</ts:RegistrationRequest >
        </soapenv:Body >
      </soapenv:Envelope > ';
    }

}