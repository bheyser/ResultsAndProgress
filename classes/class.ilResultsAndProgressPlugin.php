<?php
/* Copyright (c) 1998-2013 ILIAS open source, Extended GPL, see docs/LICENSE */
require_once 'Modules/Test/classes/class.ilTestExportPlugin.php';

/**
 * Class ilResultsAndProgressPlugin
 *
 * @author    Björn Heyser <info@bjoernheyser.de>
 * @version    $Id$
 */
class ilResultsAndProgressPlugin extends ilTestExportPlugin {
	/**
	 * Get Plugin Name.
	 * Must be same as in class name il<Name>Plugin
	 * and must correspond to plugins subdirectory name.
	 * Must be overwritten in plugin class of plugin
	 * (and should be made final)
	 *
	 * @return string Plugin Name
	 */
	function getPluginName()
	{
		return 'ResultsAndProgress';
	}
	
	/**
	 *
	 * @return string
	 */
	protected function getFormatIdentifier()
	{
		return 'rlp';
	}
	
	/**
	 *
	 * @return string
	 */
	public function getFormatLabel()
	{
		return $this->txt( 'results_and_progress_label' );
	}
	
	/**
	 *
	 * @param ilTestExportFilename $filename        	
	 */
	protected function buildExportFile(ilTestExportFilename $filename)
	{
		if( ilResultsAndProgressPlugin::isIlias51orLower() )
		{
			if( !defined('EXCEL_BACKGROUND_COLOR') ) define('EXCEL_BACKGROUND_COLOR', 'C0C0C0');
			$this->includeClass('51/PHPExcel-1.8/Classes/PHPExcel.php');
			$this->includeClass('51/class.ilExcel52x.php');
			$this->includeClass('51/class.ilAssExcelFormatHelper52x.php');
		}
		else
		{
			require_once 'Modules/TestQuestionPool/classes/class.ilAssExcelFormatHelper.php';
		}
		
		
		$this->includeClass('class.ilResultsAndProgressExportBuilder.php');
		$exportBuilder = new ilResultsAndProgressExportBuilder($this->getTest());
		$absoluteFilenameCSV = $exportBuilder->buildExportFile();
		
		if( ilResultsAndProgressPlugin::isIlias51orLower() )
		{
			$absoluteExcelFilename51 = substr($absoluteFilenameCSV, 0, -3) . 'xls';
			$absoluteExcelFilenameNEEDED = substr($absoluteFilenameCSV, 0, -3) . 'xlsx';
			rename($absoluteExcelFilename51, $absoluteExcelFilenameNEEDED);
		}
	}
	
	public static function isIlias54orGreater()
	{
		return version_compare(ILIAS_VERSION_NUMERIC, '5.4.0', '>=');
	}
	
	public static function isIlias51orLower()
	{
		return version_compare(ILIAS_VERSION_NUMERIC, '5.2.0', '<');
	}
	
	protected function phpExcelCode(ilTestExportFilename $filename)
	{
		// Creating Files with Charts using PHPExcel
		require_once './Customizing/global/plugins/Modules/Test/Export/TestStatisticsExport/classes/PHPExcel-1.8/Classes/PHPExcel.php';
		
		$objPHPExcel = new PHPExcel();
		
		// Create the first sheet with general data about the test
		$objWorksheet = $objPHPExcel->getActiveSheet();
		
		// Save XSLX file
		ilUtil::makeDirParents ( dirname ( $filename->getPathname ( 'xlsx', 'statistics' ) ) );
		$objWriter = PHPExcel_IOFactory::createWriter ( $objPHPExcel, 'Excel2007' );
		$objWriter->setIncludeCharts ( TRUE );
		$objWriter->save ( str_replace ( __FILE__, $filename->getPathname ( 'xlsx', 'statistics' ), __FILE__ ) );
	}
}