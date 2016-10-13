<?php
/**
 * Created by PhpStorm.
 * User: carlos.fernandez1
 * Date: 10/13/16
 * Time: 5:20 PM
 */
namespace McKesson;

error_reporting(E_ALL);

spl_autoload_register(function($class) {
    $source_folder = '/src/api';
    $prefix = 'McKesson';

    if (!substr($class, 0, 17) === $prefix) {
        return;
    }

    $class = substr($class, strlen($prefix));
    $location = __DIR__ . $source_folder .str_replace('\\', '/', $class) . '.php';
    if (is_file($location)) {
        require_once($location);
    }
});