<?php

/*
 * This Model will use as Data Source a XML definition and take all the needed nodes from its content.
 * @author CARLOS FERNANDEZ
 * @email carlosjfc@outlook.com
 * @company Drake College of Business
 */
App::uses('Xml', 'Utility');
App::uses('CakeSession', 'Model/Datasource');

class Widgets extends AppModel {

    public $name = 'Widgets';
    public $useTable = false;
    public $path;

    /**
     * 
     * @param int $user_id the id of the logged user 
     * @param array $logged_student the properties of a student logged into the system
     * @return array the array representation of the user rights. Later can be used to be represented as Widgets (Tiles)
     */
    public function getWidgets($user_id, $logged_student) {
        $permissions = $this->getUserPublic();
        if ($user_id != '' && !empty($user_id))
            $permissions = $this->getUserPermissions($user_id, $logged_student);
        $config = $this->getWidgetsConfig();
        $widgets_config = array();

        foreach ($permissions as $index => $user_permissions) {
            $controller = $user_permissions[0]['controller'];
            $action = $user_permissions[0]['action'];
            $group = NULL;
            $widget = $this->getWidget($config, $controller, $action, $group);
            if ($widget != NULL && !empty($widget))
                $widgets_config[$group][$controller][] = $widget;
        }

        $widget_container = array();
        foreach ($widgets_config as $igroup => $array_controllers) {
            $group = NULL;
            $setup = $this->getGroupSetup($config, $igroup);
            $group_widgets = array();
            foreach ($array_controllers as $index => $action_array) {
                $group_widgets [] = $action_array;
            }
            $widget_container[] = array(
                'id' => strtolower($igroup),
                'class' => 'tile-group',
                'style' => 'width: auto; max-width: ' . $setup['max_width'] . 'px;',
                'widgets' => $group_widgets
            );
        }


        $widget_containers = array();
        foreach ($widget_container as $key => $widget_container_array) {
            $widget_containers[] = $widget_container_array;
        }

        return $widget_containers;
    }

    /**
     * Get the setup parameters to be rendered for a given group
     * @param array $config the configuration file for the widgets
     * @param string $group the name of the group to be indexing
     * @return array|null the array with the attributes to setup the index group 
     */
    public function getGroupSetup($config, $group) {
        foreach ($config as $key_group => $group_entry) {
            foreach ($group_entry as $key_entry_group => $array_entry_group) {
                if (strtolower($key_group) == strtolower($group) && $key_entry_group == 'setup')
                    return $array_entry_group;
            }
        }
        return NULL;
    }

    /**
     * Get a set of parameters to the specific pair controler/action 
     * @param array $config the configuration file for the widgets
     * @param string $controller the name of the controller to be indexing
     * @param string $action the name of the action to be indexing
     * @param string $group the name of the group to be indexing
     * @return array|null the array with the attributes to render the represented widget for the controller/action pair
     */
    public function getWidget($config, $controller, $action, &$group) {
        foreach ($config as $key_group => $group_entry)
            foreach ($group_entry as $key_entry_group => $array_entry_group) {
                if ($key_entry_group == 'controllers') {
                    foreach ($array_entry_group as $key_controller => $controller_actions) {
                        if (strtolower($key_controller) == strtolower($controller)) {
                            foreach ($controller_actions as $key_action => $action_widget_config) {
                                if (strtolower($key_action) == strtolower($action)) {
                                    $group = $key_group;
                                    return $action_widget_config;
                                }
                            }
                        }
                    }
                }
            }
        return NULL;
    }

    /**
     * 
     * @return array the array with all the groups configuration under the widgets XML file
     */
    public function getWidgetsConfig() {
        $this->setPath();
        $load = $this->readXML('widgets', 'groups');
        return $load;
    }

    /**
     * 
     * @param array $array the array to be search
     * @param string $search the node name to start fetching
     * @return array the subarray tree founded under the $node parameter 
     */
    public function foundSubArray($array, $search) {
        $temp_array = $array;
        foreach ($temp_array as $key => $value) {
            $array = $temp_array[$key];
            if ($key == $search) {
                return $array;
            } else {
                $array = $this->foundSubArray($array, $search);
            }
        }
        return $array;
    }

    /**
     * 
     * @param string $key the XML file name to read
     * @param string|null $nodes the node name to start searching down
     * @return array array of the founded nodes under the $nodes parameter
     */
    public function readXML($key, $nodes = NULL) {
        $xml = Xml::build($this->path . $key . '.xml');
        $array = Xml::toArray($xml);
        if (!is_null($nodes)) {
            $array = $this->foundSubArray($array, $nodes);
        }
        return $array;
    }

    /**
     * 
     * @param null|string $path the path where the XML widgets configuration file will be founded
     */
    public function setPath($path = null) {
        if (is_null($path)) {
            $path = APP . 'Config' . DS;
        }
        $this->path = $path;
    }

}

?>
