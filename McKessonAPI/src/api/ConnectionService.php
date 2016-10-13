<?php
/**
 * Created by PhpStorm.
 * User: carlos.fernandez1
 * Date: 10/12/16
 * Time: 5:24 PM
 */

namespace McKesson;


class ConnectionService {
    private $soapEndPoint;
    private $soapUser;
    private $soapPassword;
    private $soapAction;
    private $soapXML;
    private $headers;

    /**
     * ConnectionService constructor.
     * @param $soapEndPoint the end point to connect and send the requests
     * @param $soapUser if there exist an user to pass along with the end point ('this one is not the same as we pass as part of the credentials on the XML Request)
     * @param $soapPassword if there exist a password to pass along with the end point ('this one is not the same as we pass as part of the credentials on the XML Request)
     * @param $soapAction the action that will be call on the end point side
     * @param $soapXML the request's Body
     */
    public function __construct($soapEndPoint, $soapUser, $soapPassword, $soapAction, $soapXML) {
        $this->soapEndPoint = $soapEndPoint;
        $this->soapUser = $soapUser;
        $this->soapPassword = $soapPassword;
        $this->soapAction = $soapAction;
        $this->soapXML = $soapXML;
        $this->headers = array("Content-type: text/xml;charset=\"utf-8\"", "Accept: text/xml", "Cache-Control: no-cache", "Pragma: no-cache", "SOAPAction: " . $soapAction, "Content-length: " . strlen($soapXML),);
    }


    /**
     * @return array|mixed return the XML that the end point give back on the communication in response to the requested SOAP Web Service action
     */
    public function request() {
        if ((strlen($this->soapXML) == 0) || ($this->soapEndPoint == '') || ($this->soapAction == '')) {
            return array();
        }

        $ch = curl_init();

        curl_setopt($ch, CURLOPT_URL, $this->soapEndPoint);
        curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, 0);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);

        if (!empty($this->soapUser) && !empty($this->soapPassword)) {
            curl_setopt($ch, CURLOPT_USERPWD, $this->soapUser . ":" . $this->soapPassword);
            // username and password
            curl_setopt($ch, CURLOPT_HTTPAUTH, CURLAUTH_ANY);
        }

        curl_setopt($ch, CURLOPT_TIMEOUT, 10);
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $this->soapXML);
        // the SOAP request
        curl_setopt($ch, CURLOPT_HTTPHEADER, $this->headers);

        $response = curl_exec($ch);
        curl_close($ch);

        return $response;
    }

    /**
     * @return string
     */
    public function getSoapEndPoint() {
        return $this->soapEndPoint;
    }

    /**
     * @param string $soapEndPoint
     */
    public function setSoapEndPoint($soapEndPoint) {
        $this->soapEndPoint = $soapEndPoint;
    }

    /**
     * @return string
     */
    public function getSoapUser() {
        return $this->soapUser;
    }

    /**
     * @param string $soapUser
     */
    public function setSoapUser($soapUser) {
        $this->soapUser = $soapUser;
    }

    /**
     * @return string
     */
    public function getSoapPassword() {
        return $this->soapPassword;
    }

    /**
     * @param string $soapPassword
     */
    public function setSoapPassword($soapPassword) {
        $this->soapPassword = $soapPassword;
    }

    /**
     * @return string
     */
    public function getSoapAction() {
        return $this->soapAction;
    }

    /**
     * @param string $soapAction
     */
    public function setSoapAction($soapAction) {
        $this->soapAction = $soapAction;
    }

    /**
     * @return string
     */
    public function getSoapXML() {
        return $this->soapXML;
    }

    /**
     * @param string $soapXML
     */
    public function setSoapXML($soapXML) {
        $this->soapXML = $soapXML;
    }


}