<?php

/* Copyright (c) 1998-2013 ILIAS open source, Extended GPL, see docs/LICENSE */

require_once 'Modules/Test/classes/class.ilTestExport.php';
require_once 'Modules/Test/classes/class.ilTestParticipantData.php';
require_once 'Services/Tracking/classes/class.ilLPStatusWrapper.php';
require_once 'Services/Tracking/classes/class.ilLearningProgressBaseGUI.php';

/**
 * Class class.ilResultsAndProgressExportBuilder
 *
 * @author    Björn Heyser <info@bjoernheyser.de>
 * @version    $Id$
 */
class ilResultsAndProgressExportBuilder extends ilTestExport
{
	const DETAIL_SHEET_OR_SHEETS_ENABLED = false;
	const EXTRA_SHEET_MAX_PARTICIPANTS = 0;
	
	/**
	 * @var ilLanguage
	 */
	protected $lang;
	
	/**
	 * @var ilTestParticipantData
	 */
	protected $participantData;
	
	/**
	 * @param ilObjTest $testObject
	 */
	public function __construct(ilObjTest $testObject)
	{
		parent::__construct($testObject, 'results');
		
		if( ilResultsAndProgressPlugin::isIlias54orGreater() )
		{
			$this->test_obj->setAccessFilteredParticipantList(
				$this->test_obj->buildStatisticsAccessFilteredParticipantList()
			);
		}
		
		$this->lang = isset($GLOBALS['DIC']) ? $GLOBALS['DIC']['lng'] : $GLOBALS['lng'];
		$db = isset($GLOBALS['DIC']) ? $GLOBALS['DIC']['ilDB'] : $GLOBALS['ilDB'];
		
		$this->participantData = new ilTestParticipantData($db, $this->lang);
		$this->participantData->load($this->test_obj->getTestId());
		
		$this->lang->loadLanguageModule('trac');
	}
	
	/**
	 * @param integer $activeId
	 * @return string
	 */
	protected function getLearningProgressByActiveId($activeId)
	{
		$usrId = $this->participantData->getUserIdByActiveId($activeId);
		$lp = ilLPStatusWrapper::_determineStatus($this->test_obj->getId(), $usrId);
		return ilLearningProgressBaseGUI::_getStatusText($lp);
	}
	
	// never used methods dealing with test object export stuff this class is never used for
	protected function initXmlExport() {}
	protected function getQuestionIds() {}
	protected function populateQuestionSetConfigXml(ilXmlWriter $xmlWriter) {}
	protected function getQuestionsQtiXml() {}
	
	
	/**
	 * Exports the evaluation data to the Microsoft Excel file format
	 *
	 * @param bool    $deliver
	 * @param string  $filterby
	 * @param string  $filtertext Filter text for the user data
	 * @param boolean $passedonly TRUE if only passed user datasets should be exported, FALSE otherwise
	 *
	 * @return string
	 */
	public function exportToExcel($deliver = TRUE, $filterby = "", $filtertext = "", $passedonly = FALSE)
	{
		$worksheet = new ilAssExcelFormatHelper();
		$worksheet->addSheet($this->lang->txt('tst_results'));
		
		$additionalFields = $this->test_obj->getEvaluationAdditionalFields();
		
		$row = 1;
		$col = 0;
		
		if($this->test_obj->getAnonymity())
		{
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('counter'));
		}
		else
		{
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('name'));
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('login'));
		}
		
		if(count($additionalFields))
		{
			foreach($additionalFields as $fieldname)
			{
				$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt($fieldname));
			}
		}
		
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_resultspoints'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('maximum_points'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_resultsmarks'));
		
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('learning_progress'));
		
		if($this->test_obj->getECTSOutput())
		{
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('ects_grade'));
		}
		
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_qworkedthrough'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_qmax'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_pworkedthrough'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_timeofwork'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_atimeofwork'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_firstvisit'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_lastvisit'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_mark_median'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_rank_participant'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_rank_median'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_total_participants'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('tst_stat_result_median'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('scored_pass'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('pass'));
		
		$worksheet->setBold('A' . $row . ':' . $worksheet->getColumnCoord($col - 1) . $row);
		
		$counter = 1;
		$data = $this->test_obj->getCompleteEvaluationData(TRUE, $filterby, $filtertext);
		$firstrowwritten = false;
		foreach($data->getParticipants() as $active_id => $userdata)
		{
			if($passedonly && $data->getParticipant($active_id)->getPassed() == FALSE)
			{
				continue;
			}
			
			$row++;
			$col = 0;
			
			// each participant gets an own row for question column headers
			if($this->test_obj->isRandomTest())
			{
				$row++;
			}
			
			if($this->test_obj->getAnonymity())
			{
				$worksheet->setCell($row, $col++, $counter);
			}
			else
			{
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getName());
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getLogin());
			}
			
			if(count($additionalFields))
			{
				$userfields = ilObjUser::_lookupFields($userdata->getUserId());
				foreach ($additionalFields as $fieldname)
				{
					if(strcmp($fieldname, 'gender') == 0)
					{
						$worksheet->setCell($row, $col++, $this->lang->txt('gender_' . $userfields[$fieldname]));
					}
					else
					{
						$worksheet->setCell($row, $col++, $userfields[$fieldname]);
					}
				}
			}
			
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getReached());
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getMaxpoints());
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getMark());
			
			$worksheet->setCell($row, $col++, $this->getLearningProgressByActiveId($active_id));
			
			if($this->test_obj->getECTSOutput())
			{
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getECTSMark());
			}
			
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getQuestionsWorkedThrough());
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getNumberOfQuestions());
			$worksheet->setCell($row, $col++, ($data->getParticipant($active_id)->getQuestionsWorkedThroughInPercent() / 100.0) . '%');
			
			$time = $data->getParticipant($active_id)->getTimeOfWork();
			$time_seconds = $time;
			$time_hours    = floor($time_seconds/3600);
			$time_seconds -= $time_hours   * 3600;
			$time_minutes  = floor($time_seconds/60);
			$time_seconds -= $time_minutes * 60;
			$worksheet->setCell($row, $col++, sprintf("%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds));
			$time = $data->getParticipant($active_id)->getQuestionsWorkedThrough() ? $data->getParticipant($active_id)->getTimeOfWork() / $data->getParticipant($active_id)->getQuestionsWorkedThrough() : 0;
			$time_seconds = $time;
			$time_hours    = floor($time_seconds/3600);
			$time_seconds -= $time_hours   * 3600;
			$time_minutes  = floor($time_seconds/60);
			$time_seconds -= $time_minutes * 60;
			$worksheet->setCell($row, $col++, sprintf("%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds));
			$worksheet->setCell($row, $col++, new ilDateTime($data->getParticipant($active_id)->getFirstVisit(), IL_CAL_UNIX));
			$worksheet->setCell($row, $col++, new ilDateTime($data->getParticipant($active_id)->getLastVisit(), IL_CAL_UNIX));
			
			$median = $data->getStatistics()->getStatistics()->median();
			$pct = $data->getParticipant($active_id)->getMaxpoints() ? $median / $data->getParticipant($active_id)->getMaxpoints() * 100.0 : 0;
			$mark = $this->test_obj->mark_schema->getMatchingMark($pct);
			$mark_short_name = "";
			
			if(is_object($mark))
			{
				$mark_short_name = $mark->getShortName();
			}
			
			$worksheet->setCell($row, $col++, $mark_short_name);
			$worksheet->setCell($row, $col++, $data->getStatistics()->getStatistics()->rank($data->getParticipant($active_id)->getReached()));
			$worksheet->setCell($row, $col++, $data->getStatistics()->getStatistics()->rank_median());
			$worksheet->setCell($row, $col++, $data->getStatistics()->getStatistics()->count());
			$worksheet->setCell($row, $col++, $median);
			
			if($this->test_obj->getPassScoring() == SCORE_BEST_PASS)
			{
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getBestPass() + 1);
			}
			else
			{
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getLastPass() + 1);
			}
			
			$startcol = $col;
			
			for($pass = 0; $pass <= $data->getParticipant($active_id)->getLastPass(); $pass++)
			{
				$col = $startcol;
				$finishdate = $this->test_obj->getPassFinishDate($active_id, $pass);
				if($finishdate > 0)
				{
					if ($pass > 0)
					{
						$row++;
						if ($this->test_obj->isRandomTest())
						{
							$row++;
						}
					}
					$worksheet->setCell($row, $col++, $pass + 1);
					if(is_object($data->getParticipant($active_id)) && is_array($data->getParticipant($active_id)->getQuestions($pass)))
					{
						$evaluatedQuestions = $data->getParticipant($active_id)->getQuestions($pass);
						
						if( $this->test_obj->getShuffleQuestions() )
						{
							// reorder questions according to general fixed sequence,
							// so participant rows can share single questions header
							$questions = array();
							foreach($this->test_obj->getQuestions() as $qId)
							{
								foreach($evaluatedQuestions as $evaledQst)
								{
									if( $evaledQst['id'] != $qId )
									{
										continue;
									}
									
									$questions[] = $evaledQst;
								}
							}
						}
						else
						{
							$questions = $evaluatedQuestions;
						}
						
						foreach($questions as $question)
						{
							$question_data = $data->getParticipant($active_id)->getPass($pass)->getAnsweredQuestionByQuestionId($question["id"]);
							$worksheet->setCell($row, $col, $question_data["reached"]);
							if($this->test_obj->isRandomTest())
							{
								// random test requires question headers for every participant
								// and we allready skipped a row for that reason ( --> row - 1)
								$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col) . ($row - 1),  preg_replace("/<.*?>/", "", $data->getQuestionTitle($question["id"])));
							}
							else
							{
								if($pass == 0 && !$firstrowwritten)
								{
									$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col) . 1, $data->getQuestionTitle($question["id"]));
								}
							}
							$col++;
						}
						$firstrowwritten = true;
					}
				}
			}
			$counter++;
		}
		
		if(self::DETAIL_SHEET_OR_SHEETS_ENABLED)
		if($this->test_obj->getExportSettingsSingleChoiceShort() && !$this->test_obj->isRandomTest() && $this->test_obj->hasSingleChoiceQuestions())
		{
			// special tab for single choice tests
			$titles = $this->test_obj->getQuestionTitlesAndIndexes();
			$positions = array();
			$pos = 0;
			$row = 1;
			foreach($titles as $id => $title)
			{
				$positions[$id] = $pos;
				$pos++;
			}
			
			$usernames = array();
			$participantcount = count($data->getParticipants());
			$allusersheet = false;
			$pages = 0;
			
			$worksheet->addSheet($this->lang->txt('eval_all_users'));
			
			$col = 0;
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lang->txt('name'));
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row,  $this->lang->txt('login'));
			if(count($additionalFields))
			{
				foreach($additionalFields as $fieldname)
				{
					if(strcmp($fieldname, "matriculation") == 0)
					{
						$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row,  $this->lang->txt('matriculation'));
					}
				}
			}
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row,  $this->lang->txt('test'));
			foreach($titles as $title)
			{
				$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $title);
			}
			$worksheet->setBold('A' . $row . ':' . $worksheet->getColumnCoord($col - 1) . $row);
			
			$row++;
			foreach($data->getParticipants() as $active_id => $userdata)
			{
				$username = (!is_null($userdata) && $userdata->getName()) ? $userdata->getName() : "ID $active_id";
				if (array_key_exists($username, $usernames))
				{
					$usernames[$username]++;
					$username .= " ($usernames[$username])";
				}
				else
				{
					$usernames[$username] = 1;
				}
				$col = 0;
				$worksheet->setCell($row, $col++, $username);
				$worksheet->setCell($row, $col++, $userdata->getLogin());
				if (count($additionalFields))
				{
					$userfields = ilObjUser::_lookupFields($userdata->getUserID());
					foreach ($additionalFields as $fieldname)
					{
						if (strcmp($fieldname, "matriculation") == 0)
						{
							if (strlen($userfields[$fieldname]))
							{
								$worksheet->setCell($row, $col++, $userfields[$fieldname]);
							}
							else
							{
								$col++;
							}
						}
					}
				}
				$worksheet->setCell($row, $col++, $this->test_obj->getTitle());
				$pass = $userdata->getScoredPass();
				if(is_object($userdata) && is_array($userdata->getQuestions($pass)))
				{
					foreach($userdata->getQuestions($pass) as $question)
					{
						$objQuestion = assQuestion::_instantiateQuestion($question["id"]);
						if(is_object($objQuestion) && strcmp($objQuestion->getQuestionType(), 'assSingleChoice') == 0)
						{
							$solution = $objQuestion->getSolutionValues($active_id, $pass);
							$pos = $positions[$question["id"]];
							$selectedanswer = "x";
							foreach ($objQuestion->getAnswers() as $id => $answer)
							{
								if (strlen($solution[0]["value1"]) && $id == $solution[0]["value1"])
								{
									$selectedanswer = $answer->getAnswertext();
								}
							}
							$worksheet->setCell($row, $col+$pos, $selectedanswer);
						}
					}
				}
				$row++;
			}
			
			if($this->test_obj->isSingleChoiceTestWithoutShuffle())
			{
				// special tab for single choice tests without shuffle option
				$pos = 0;
				$row = 1;
				$usernames = array();
				$allusersheet = false;
				$pages = 0;
				
				$worksheet->addSheet($this->lang->txt('eval_all_users'). ' (2)');
				
				$col = 0;
				$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row,  $this->lang->txt('name'));
				$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row,  $this->lang->txt('login'));
				if (count($additionalFields))
				{
					foreach ($additionalFields as $fieldname)
					{
						if (strcmp($fieldname, "matriculation") == 0)
						{
							$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row,  $this->lang->txt('matriculation'));
						}
					}
				}
				$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row,  $this->lang->txt('test'));
				foreach($titles as $title)
				{
					$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row,  $title);
				}
				$worksheet->setBold('A' . $row . ':' . $worksheet->getColumnCoord($col - 1) . $row);
				
				$row++;
				foreach ($data->getParticipants() as $active_id => $userdata)
				{
					$username = (!is_null($userdata) && $userdata->getName()) ? $userdata->getName() : "ID $active_id";
					if (array_key_exists($username, $usernames))
					{
						$usernames[$username]++;
						$username .= " ($usernames[$username])";
					}
					else
					{
						$usernames[$username] = 1;
					}
					$col = 0;
					$worksheet->setCell($row, $col++, $username);
					$worksheet->setCell($row, $col++, $userdata->getLogin());
					if (count($additionalFields))
					{
						$userfields = ilObjUser::_lookupFields($userdata->getUserId());
						foreach ($additionalFields as $fieldname)
						{
							if (strcmp($fieldname, "matriculation") == 0)
							{
								if (strlen($userfields[$fieldname]))
								{
									$worksheet->setCell($row, $col++, $userfields[$fieldname]);
								}
								else
								{
									$col++;
								}
							}
						}
					}
					$worksheet->setCell($row, $col++, $this->test_obj->getTitle());
					$pass = $userdata->getScoredPass();
					if(is_object($userdata) && is_array($userdata->getQuestions($pass)))
					{
						foreach($userdata->getQuestions($pass) as $question)
						{
							$objQuestion = ilObjTest::_instanciateQuestion($question["aid"]);
							if(is_object($objQuestion) && strcmp($objQuestion->getQuestionType(), 'assSingleChoice') == 0)
							{
								$solution = $objQuestion->getSolutionValues($active_id, $pass);
								$pos = $positions[$question["aid"]];
								$selectedanswer = chr(65+$solution[0]["value1"]);
								$worksheet->setCell($row, $col+$pos, $selectedanswer);
							}
						}
					}
					$row++;
				}
			}
		}
		else
		{
			// test participant result export
			$usernames = array();
			$participantcount = count($data->getParticipants());
			$allusersheet = false;
			$pages = 0;
			$i = 0;
			foreach($data->getParticipants() as $active_id => $userdata)
			{
				$i++;
				
				$username = (!is_null($userdata) && $userdata->getName()) ? $userdata->getName() : "ID $active_id";
				if(array_key_exists($username, $usernames))
				{
					$usernames[$username]++;
					$username .= " ($i)";
				}
				else
				{
					$usernames[$username] = 1;
				}
				
				if($participantcount > EXTRA_SHEET_MAX_PARTICIPANTS)
				{
					if(!$allusersheet || ($pages-1) < floor($row / 64000))
					{
						$worksheet->addSheet($this->lang->txt("eval_all_users") . (($pages > 0) ? " (".($pages+1).")" : ""));
						$allusersheet = true;
						$row = 1;
						$pages++;
					}
				}
				else
				{
					$resultsheet = $worksheet->addSheet($username);
				}
				
				$pass = $userdata->getScoredPass();
				$row = ($allusersheet) ? $row : 1;
				$worksheet->setCell($row, 0, sprintf($this->lang->txt("tst_result_user_name_pass"), $pass+1, $userdata->getName()));
				$worksheet->setBold($worksheet->getColumnCoord(0) . $row);
				$row += 2;
				if(is_object($userdata) && is_array($userdata->getQuestions($pass)))
				{
					foreach($userdata->getQuestions($pass) as $question)
					{
						require_once "./Modules/TestQuestionPool/classes/class.assQuestion.php";
						$question = assQuestion::_instanciateQuestion($question["id"]);
						if(is_object($question))
						{
							$row = $question->setExportDetailsXLS($worksheet, $row, $active_id, $pass);
						}
					}
				}
			}
		}
		
		if($deliver)
		{
			$testname = $this->test_obj->getTitle();
			switch($this->mode)
			{
				case 'results':
					$testname .= '_results';
					break;
			}
			$testname = ilUtil::getASCIIFilename(preg_replace("/\s/", "_", $testname)) . '.xlsx';
			$worksheet->sendToClient($testname);
		}
		else
		{
			$excelfile = ilUtil::ilTempnam();
			$worksheet->writeToFile($excelfile);
			return $excelfile . '.xlsx';
		}
	}
	
	/**
	 * Exports the evaluation data to the CSV file format
	 *
	 * @param string $filtertext Filter text for the user data
	 * @param boolean $passedonly TRUE if only passed user datasets should be exported, FALSE otherwise
	 * @access public
	 */
	function exportToCSV($deliver = TRUE, $filterby = "", $filtertext = "", $passedonly = FALSE)
	{
		$rows = array();
		$datarow = array();
		$col = 1;
		if ($this->test_obj->getAnonymity())
		{
			array_push($datarow, $this->lang->txt("counter"));
			$col++;
		}
		else
		{
			array_push($datarow, $this->lang->txt("name"));
			$col++;
			array_push($datarow, $this->lang->txt("login"));
			$col++;
		}
		$additionalFields = $this->test_obj->getEvaluationAdditionalFields();
		if (count($additionalFields))
		{
			foreach ($additionalFields as $fieldname)
			{
				array_push($datarow, $this->lang->txt($fieldname));
				$col++;
			}
		}
		array_push($datarow, $this->lang->txt("tst_stat_result_resultspoints"));
		$col++;
		array_push($datarow, $this->lang->txt("maximum_points"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_resultsmarks"));
		$col++;
		
		array_push($datarow, $this->lang->txt("learning_progress"));
		$col++;
		
		if ($this->test_obj->getECTSOutput())
		{
			array_push($datarow, $this->lang->txt("ects_grade"));
			$col++;
		}
		array_push($datarow, $this->lang->txt("tst_stat_result_qworkedthrough"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_qmax"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_pworkedthrough"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_timeofwork"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_atimeofwork"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_firstvisit"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_lastvisit"));
		$col++;
		
		array_push($datarow, $this->lang->txt("tst_stat_result_mark_median"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_rank_participant"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_rank_median"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_total_participants"));
		$col++;
		array_push($datarow, $this->lang->txt("tst_stat_result_median"));
		$col++;
		array_push($datarow, $this->lang->txt("scored_pass"));
		$col++;
		
		array_push($datarow, $this->lang->txt("pass"));
		$col++;
		
		$data =& $this->test_obj->getCompleteEvaluationData(TRUE, $filterby, $filtertext);
		$headerrow = $datarow;
		$counter = 1;
		foreach ($data->getParticipants() as $active_id => $userdata)
		{
			$datarow = $headerrow;
			$remove = FALSE;
			if ($passedonly)
			{
				if ($data->getParticipant($active_id)->getPassed() == FALSE)
				{
					$remove = TRUE;
				}
			}
			if (!$remove)
			{
				$datarow2 = array();
				if ($this->test_obj->getAnonymity())
				{
					array_push($datarow2, $counter);
				}
				else
				{
					array_push($datarow2, $data->getParticipant($active_id)->getName());
					array_push($datarow2, $data->getParticipant($active_id)->getLogin());
				}
				if (count($additionalFields))
				{
					$userfields = ilObjUser::_lookupFields($userdata->getUserID());
					foreach ($additionalFields as $fieldname)
					{
						if (strcmp($fieldname, "gender") == 0)
						{
							array_push($datarow2, $this->lang->txt("gender_" . $userfields[$fieldname]));
						}
						else
						{
							array_push($datarow2, $userfields[$fieldname]);
						}
					}
				}
				array_push($datarow2, $data->getParticipant($active_id)->getReached());
				array_push($datarow2, $data->getParticipant($active_id)->getMaxpoints());
				array_push($datarow2, $data->getParticipant($active_id)->getMark());
				
				array_push($datarow2, $this->getLearningProgressByActiveId($active_id));
				
				if ($this->test_obj->getECTSOutput())
				{
					array_push($datarow2, $data->getParticipant($active_id)->getECTSMark());
				}
				array_push($datarow2, $data->getParticipant($active_id)->getQuestionsWorkedThrough());
				array_push($datarow2, $data->getParticipant($active_id)->getNumberOfQuestions());
				array_push($datarow2, $data->getParticipant($active_id)->getQuestionsWorkedThroughInPercent() / 100.0);
				$time = $data->getParticipant($active_id)->getTimeOfWork();
				$time_seconds = $time;
				$time_hours    = floor($time_seconds/3600);
				$time_seconds -= $time_hours   * 3600;
				$time_minutes  = floor($time_seconds/60);
				$time_seconds -= $time_minutes * 60;
				array_push($datarow2, sprintf("%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds));
				$time = $data->getParticipant($active_id)->getQuestionsWorkedThrough() ? $data->getParticipant($active_id)->getTimeOfWork() / $data->getParticipant($active_id)->getQuestionsWorkedThrough() : 0;
				$time_seconds = $time;
				$time_hours    = floor($time_seconds/3600);
				$time_seconds -= $time_hours   * 3600;
				$time_minutes  = floor($time_seconds/60);
				$time_seconds -= $time_minutes * 60;
				array_push($datarow2, sprintf("%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds));
				
				$fv = $data->getParticipant($active_id)->getFirstVisit();
				$lv = $data->getParticipant($active_id)->getLastVisit();
				foreach(array($fv, $lv) as $ts)
				{
					if($ts)
					{
						$visit = ilDatePresentation::formatDate(new ilDateTime($ts, IL_CAL_UNIX));
						array_push($datarow2, $visit);
					}
					else
					{
						array_push($datarow2, "");
					}
				}
				
				$median = $data->getStatistics()->getStatistics()->median();
				$pct = $data->getParticipant($active_id)->getMaxpoints() ? $median / $data->getParticipant($active_id)->getMaxpoints() * 100.0 : 0;
				$mark = $this->test_obj->mark_schema->getMatchingMark($pct);
				$mark_short_name = "";
				if (is_object($mark))
				{
					$mark_short_name = $mark->getShortName();
				}
				array_push($datarow2, $mark_short_name);
				array_push($datarow2, $data->getStatistics()->getStatistics()->rank($data->getParticipant($active_id)->getReached()));
				array_push($datarow2, $data->getStatistics()->getStatistics()->rank_median());
				array_push($datarow2, $data->getStatistics()->getStatistics()->count());
				array_push($datarow2, $median);
				if ($this->test_obj->getPassScoring() == SCORE_BEST_PASS)
				{
					array_push($datarow2, $data->getParticipant($active_id)->getBestPass() + 1);
				}
				else
				{
					array_push($datarow2, $data->getParticipant($active_id)->getLastPass() + 1);
				}
				for ($pass = 0; $pass <= $data->getParticipant($active_id)->getLastPass(); $pass++)
				{
					$finishdate = $this->test_obj->getPassFinishDate($active_id, $pass);
					if ($finishdate > 0)
					{
						if ($pass > 0)
						{
							for ($i = 1; $i < $col-1; $i++)
							{
								array_push($datarow2, "");
								array_push($datarow, "");
							}
							array_push($datarow, "");
						}
						array_push($datarow2, $pass+1);
						if (is_object($data->getParticipant($active_id)) && is_array($data->getParticipant($active_id)->getQuestions($pass)))
						{
							foreach ($data->getParticipant($active_id)->getQuestions($pass) as $question)
							{
								$question_data = $data->getParticipant($active_id)->getPass($pass)->getAnsweredQuestionByQuestionId($question["id"]);
								array_push($datarow2, $question_data["reached"]);
								array_push($datarow, preg_replace("/<.*?>/", "", $data->getQuestionTitle($question["id"])));
							}
						}
						if ($this->test_obj->isRandomTest() || $this->test_obj->getShuffleQuestions() || ($counter == 1 && $pass == 0))
						{
							array_push($rows, $datarow);
						}
						$datarow = array();
						array_push($rows, $datarow2);
						$datarow2 = array();
					}
				}
				$counter++;
			}
		}
		$csv = "";
		$separator = ";";
		foreach ($rows as $evalrow)
		{
			$csvrow =& $this->test_obj->processCSVRow($evalrow, TRUE, $separator);
			$csv .= join($csvrow, $separator) . "\n";
		}
		if ($deliver)
		{
			ilUtil::deliverData($csv, ilUtil::getASCIIFilename($this->test_obj->getTitle() . "_results.csv"));
			exit;
		}
		else
		{
			return $csv;
		}
	}
	
	
}