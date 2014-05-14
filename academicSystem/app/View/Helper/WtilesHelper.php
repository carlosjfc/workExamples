<?php

/**
 * This helper will be doing all the heavy work to render the read widgets XML configuration file
 * into the proper HTML5 elements 
 * @author CARLOS FERNANDEZ
 * @email carlosjfc@outlook.com
 * @company Drake College of Business
 */

/**
 * CakePHP WtilesHelper: This CakePHP Helper will allow to take a Tile definition as discused in http://metroui.org.ua/ and create the tiles groups
 * @author Carlos Fernandez
 */
class WtilesHelper extends AppHelper {

    public $helpers = array('Html');

    public function __construct(View $View, $settings = array()) {
        parent::__construct($View, $settings);
    }

    /**
     * @param array $widget_containers the nested array containing the Group of Tiles to be rendered
     * @return string the HTML representation of the nested widgets
     */
    public function createWidgets($widget_containers) {
        $groups_html = '';
        foreach ($widget_containers as $index => $values) {
            $id = $values['id'];
            $class = $values['class'];
            $style = $values['style'];
            $widgets = $values['widgets'];
            $widgets_html = '';
            foreach ($widgets as $indice => $widget_group)
                foreach ($widget_group as $key => $widget_tile)
                    $widgets_html .= $this->contentGenerator($widget_tile);

            $groups_html .= $this->Html->div(null, $widgets_html, array('id' => $id, 'class' => $class, 'style' => $style));
        }
        return $groups_html;
    }

    /**
     * A tile can have many kind of HTML elements, while iterating in the many nested elements this function will iterate recursevly
     * until no nodes where analized and represented in its HTML syntaxys
     * @param array $element_array 
     * @return string the HTML representation of the tile
     */
    public function contentGenerator($element_array) {
        $element_id = '';
        $element_class = '';
        $element_tag = '';
        $element_attributes_array = '';
        $element_url = '';
        $element_internalurl = '';
        $element_ajax_content = '';
        $element_content = '';
        $element_options = array();
        $element_content_html = '';
        $element_html = '';

        if (isset($element_array['id']) && !empty($element_array['id'])) {
            $element_id = $element_array['id'];
            $element_options = array('id' => $element_id);
        }

        if (isset($element_array['class']) && !empty($element_array['class'])) {
            $element_class_array = $element_array['class'];
            foreach ($element_class_array as $key => $value) {
                if ($value != '')
                    $element_class .= $value . ' ';
            }
            $element_class = trim($element_class);
            $element_options += array('class' => $element_class);
        }

        if (isset($element_array['tag']) && !empty($element_array['tag'])) {
            $element_tag = $element_array['tag'];
        }

        if (isset($element_array['attributes']) && !empty($element_array['attributes'])) {
            $element_attributes_array = $element_array['attributes'];
            $element_options += $element_attributes_array;
        }

        if (isset($element_array['url']) && !empty($element_array['url'])) {
            $element_url_array = $element_array['url'];
            $element_url_parameters = array();
            foreach ($element_url_array as $url_key => $url_value) {
                if ($url_key == 'parameters' && !empty($url_value)) {
                    foreach ($url_value as $key => $value) {
                        if ($value != '' && !empty($value))
                            $element_url_parameters[] = $value;
                    }
                }
            }
            if (!empty($element_url_array['controller'])) {
                $element_url = array('controller' => $element_url_array['controller'], 'action' => $element_url_array['action']);
                if (!empty($element_url_parameters))
                    $element_url = array_merge($element_url, $element_url_parameters);
            }
        }

        if (isset($element_array['ajax_content']) && !empty($element_array['ajax_content'])) {
            $element_ajax_content = $this->contentGenerator($element_array['ajax_content']);
        }

        if (isset($element_array['internalurl']) && !empty($element_array['internalurl'])) {
            $element_internalurl = $this->contentGenerator($element_array['internalurl']);
        }

        if (isset($element_array['content']) && !empty($element_array['content'])) {
            $element_content = $element_array['content'];
        }
        if (is_array($element_content)) {
            foreach ($element_content as $key => $array_of_contents) {
                $element_content_html .= $this->contentGenerator($array_of_contents);
            }
        } else {
            $element_content_html = $element_content;
        }

        if ($element_tag == 'div') {
            $element_html = $this->Html->div(null, $element_content_html . $element_internalurl . $element_ajax_content, $element_options);
        } elseif ($element_tag == 'a') {
            $element_html = $this->Html->link($element_content_html . $element_internalurl . $element_ajax_content, $element_url, $element_options);
        } elseif ($element_tag == 'img') {
            $element_html = $this->Html->image($element_content_html . $element_internalurl . $element_ajax_content, $element_options);
            if (!empty($element_url)) {
                $element_url = array('url' => $element_url);
                $element_html = $this->Html->image($element_content_html . $element_internalurl . $element_ajax_content, array_merge($element_options, $element_url));
            }
        } elseif ($element_tag == 'p') {
            $element_html = $this->Html->para(null, $element_content_html . $element_internalurl . $element_ajax_content, $element_options);
        } else {
            $element_html = $this->Html->tag($element_tag, $element_content_html . $element_internalurl . $element_ajax_content, $element_options);
        }

        return $element_html;
    }

}