<?php
/**
 * Created by PhpStorm.
 * User: carlos.fernandez1
 * Date: 10/13/16
 * Time: 12:19 PM
 */

namespace McKesson;


class Utils {
    /**
     * Load the constants defined in a configuration file into Memory to be use later for the application.
     * @param array $config The results of parsing the configuration file.
     * @static
     */
    static public function loadConsts($config) {
        $configParsed = Utils::parseConfig($config);
        $constants = Utils::getConstantsKeys($configParsed);
        if (!empty($constants)) {
            foreach ($constants as $constant) {
                if (!defined($constant) && !empty($configParsed[$constant])) {
                    define($constant, $configParsed[$constant]);
                }
            }
        }
    }

    static private function parseConfig($configFile) {
        return parse_ini_file($configFile, FALSE, INI_SCANNER_RAW);
    }

    /**
     * Return the array's keys on the $config file. They are the constant names to be loaded into the system.
     * @param array $config The results of parsing the configuration file.
     * @return array The constants names because they are the Key on the $conf array.
     * @static
     */
    static private function getConstantsKeys($config) {
        $constants = array();
        foreach ($config as $key => $value) {
            if ($key === strtoupper($key)) {
                $constants[] = $key;
            }
        }
        return $constants;
    }
}