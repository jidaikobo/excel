<?php
namespace Excel;

trait Trait_Method
{
	/*
	 * $ln 
	 * R: 右へ移動(既定)
	 * D: 下へ移動
	 * U: 上に移動
	 * L: 左へ移動
	 */
	public function Box($data, $ln = 'R')
	{
		$location = $this->getXY();

		$sheet = static::$_sheet;

		$sheet->setCellValue($location, $data);
		switch ($ln)
		{
			case 'U':
			case 'T':
				return $this->upWard();
				break;
			case 'D':
			case 'B':
				return $this->downWard();
				break;
			case 'R':
				return $this->rightWard();
				break;
			case 'L':
				return $this->leftWard();
				break;
			default:
				return false;
				break;
		}
	}

	public function upWard($distance = 1)
	{
		$y = static::$_current_y;
		$y = $y - $distance;
		if ($y < 1) $y = 1;
		static::$_current_y = $y;
	}

	public function downWard($distance = 1)
	{
		static::$_current_y = static::$_current_y + $distance;
	}

	public function leftWard($distance = 1)
	{
		$x = static::$_current_x;
		for ($i=0; $i<$distance; $i++)
		{
			$x--;
		}
		if ($x < 'A') $x = $x;

		static::$_current_x = $x;
	}

	public function rightWard($distance = 1)
	{
		$x = static::$_current_x;
		for ($i=0; $i<$distance; $i++)
		{
			$x++;
		}

		static::$_current_x = $x;
	}

	public function setStyle($cell, $styles)
	{
	}
}
