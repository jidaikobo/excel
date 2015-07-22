<?php
/**
 * Excel Package
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

	use \Excel\Trait_Method;
	use \Excel\Trait_Wrapper;

	private static $_current_x = 'A';
	private static $_current_y = '1';
	private static $_sheet = null;

	public function getX()
	{
		return static::$_current_x;
	}

	public function getY()
	{
		return static::$_current_y;
	}

	public function getXY()
	{
		return static::$_current_x .  static::$_current_y;
	}

	public function setX($value = 'A')
	{
		if (preg_match('/^[A-Z]+$/', $value))
		{
			static::$_current_x = $value;
			return static::$_current_x;
		}
		else if (is_int($value) && $value>0)
		{
			// 1 からスタートなので -1
			$value--;
			static::$_current_x = $this->alpha_incremeter($value);
			return static::$_current_x;
		}
		else
		{
			return false; // todo throw error or return current value
		}
	}

	private function alpha_incremeter($value, $alpha = 'A')
	{
		for ($i=0; $i<$value; $i++)
		{
			$alpha++;
		}
		return $alpha;
		/*
		$add = $value%26;
		$chr = chr((ord($alpha)) + $add);

		$multi = floor($value/26);
		if ($multi)
		{
			$multi--;
			return $chr = $this->alpha_incremeter($multi, 'A') . $chr;
		}
		else {
			return $chr;
		}*/
	}

	/*
	public function alpha2num($alpha)
	{
		preg_match_all('/[A-Z]/', $alpha, $matches);

		if (count($matches) < 1) return false;

		$alphas = array_reverse($matches[0]);
		$retval = 0;
		foreach ($alphas as $digit => $val)
		{
			$val = ord($val) - ord('@');
			if ($digit == 0) $retval += $val;
			$retval += pow(26, $digit)*$val;
		}

		return $retval;
	}

	public function num2alpha($alpha)
	{
		preg_match_all('/[A-Z]/', $alpha, $matches);

		if (count($matches) < 1) return false;

		$alphas = array_reverse($matches[0]);
		$retval = 0;
		foreach ($alphas as $digit => $val)
		{
			$val = ord($val) - ord('@');
			if ($digit == 0) $retval += $val;
			$retval += $digit*26*$val;
		}

		return $retval;
	}
	 */


	public function setY($value = 'A')
	{
		if (is_int($value) && $value>0)
		{
			static::$_current_y = $value;
			return static::$_current_y;
		}
		else
		{
			return false; // todo throw error or return current value
		}
	}

	public function setXY($x, $y)
	{
		$retX = $this->setX($x);
		$retY = $this->setY($y);

		return $retX . $retY;
	}

	public static function forge()
	{
		return new static();
	}

	public function output($filename = 'sample', $book = null, $format = 'Excel2007')
	{
		if (!$book) $book = $this;

		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
		header('Cache-Control: max-age=0');
		// If you're serving to IE 9, then the following may be needed
		// header('Cache-Control: max-age=1');

		// If you're serving to IE over SSL, then the following may be needed
		header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
		header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
		header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
		header ('Pragma: public'); // HTTP/1.0

		$objWriter = \PHPExcel_IOFactory::createWriter($book, $format);
		$objWriter->save('php://output');
		exit();
	}


	public function __get($name)
	{
		if ($name == 'sheet') {
			if (is_null(static::$_sheet)) {
				static::$_sheet = $this->getActiveSheet();
			}
			return static::$_sheet;
		}
	}

	public function setSheet($sheet)
	{
		static::$_sheet = $sheet;
		return static::$_sheet;
	}
}
