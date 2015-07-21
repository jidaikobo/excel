<?php
/**
 * Pdf Package
 *
 * @package		Excel
 * @author		konagai@jidaikobo.com
 * @copyright	Copyright (c) jidaikobo Inc.
 * @license		Public Domain
 * @link		http://www.jidaikobo.com/
 */
namespace Excel;

include_once( dirname(dirname( __FILE__ )) . '/lib/PHPExcel/Classes/PHPExcel.php');
include_once( dirname(dirname( __FILE__ )) . '/lib/PHPExcel/Classes/PHPExcel/IOFactory.php' );

class Excel extends \PHPExcel
{

	public static function forge()
	{
		return new static();
	}

	public function output($filename = 'sample.xls', $format = 'Excel5')
	{
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="' . $filename . '"');
		header('Cache-Control: max-age=0');
		// If you're serving to IE 9, then the following may be needed
		// header('Cache-Control: max-age=1');

		// If you're serving to IE over SSL, then the following may be needed
		header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
		header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
		header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
		header ('Pragma: public'); // HTTP/1.0

		$objWriter = \PHPExcel_IOFactory::createWriter($this, $format);
		$objWriter->save('php://output');
		exit();
	}
}
