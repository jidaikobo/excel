<?php
/**
 */

Autoloader::add_core_namespace('Pdf');

Autoloader::add_classes(array(
	'Excel\\Excel'          => __DIR__ . '/classes/excel.php',
	'Excel\\Trait_Wrapper'  => __DIR__ . '/classes/trait/wrapper.php',
	'Excel\\Trait_Method'   => __DIR__ . '/classes/trait/method.php',
));
