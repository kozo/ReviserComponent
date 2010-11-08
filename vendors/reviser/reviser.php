<?PHP
/**
 * Excel_Reviser Version 0.33beta  Author:kishiyan
 *		with Image OBJ patch       Co-author:sake&ume
 * Copyright (c) 2006-2009 kishiyan <excelreviser@gmail.com>
 * All rights reserved.
 *
 *   URL  http://chazuke.com/forum/viewforum.php?f=3
 *
 * Redistribution and use in source, with or without modification, are
 * permitted provided that the following conditions are met:
 * 1. Redistributions of source code must retain the above copyright
 *    notice, this list of conditions and the following disclaimer,
 *    without modification, immediately at the beginning of the file.
 * 2. The name of the author may not be used to endorse or promote products
 *    derived from this software without specific prior written permission.
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *   URL http://www.gnu.org/licenses/gpl.html
 * 
 * @package Excel_Reviser
 * @author kishiyan <excelreviser@gmail.com>
 * @copyright Copyright &copy; 2006-2009, kishiyan
 * @since PHP 4.4.1 w/mbstring,GD
 * @version 0.33 beta 2009/02/13
 */

/*  HISTORY
	refer README_EUC.txt
*/

define('Reviser_Version','0.33alpha');
define('Version_Num', 0.33);

define('Default_CHARSET', 'eucJP-win');
define('Code_BIFF8', 0x600);
define('Code_WorkbookGlobals', 0x5);
define('Code_Worksheet', 0x10);
define('Type_EOF', 0x0a);
define('Type_BOUNDSHEET', 0x85);
define('Type_SST', 0xfc);
define('Type_CONTINUE', 0x3c);
define('Type_EXTSST', 0xff);
define('Type_LABEL', 0x204);
define('Type_LABELSST', 0xfd);
define('Type_WRITEACCESS', 0x5c);
define('Type_OBJPROJ', 0xd3);
define('Type_BUTTONPROPERTYSET', 0x1ba);
define('Type_DIMENSION', 0x200);
define('Type_ROW', 0x208);
define('Type_DEFCOLWIDTH', 0x55);
define('Type_COLINFO', 0x7d);
define('Type_DBCELL', 0xd7);
define('Type_RK', 0x7e);
define('Type_RK2', 0x27e);
define('Type_MULRK', 0xbd);
define('Type_MULBLANK', 0xbe);
define('Type_INDEX', 0x20b);
define('Type_NUMBER', 0x203);
define('Type_FORMULA', 0x406);
define('Type_FORMULA2', 0x6);
define('Type_BOOLERR', 0x205);
define('Type_UNKNOWN', 0xffff);
define('Type_BLANK', 0x201);
define('Type_SharedFormula', 0x4bc);
define('Type_STRING', 0x207);
define('Type_HEADER', 0x14);
define('Type_FOOTER', 0x15);
define('Type_BOF', 0x809);
define('Type_WINDOW2', 0x23e);
define('Type_COUNTRY', 0x8c);
define('Type_SUPBOOK', 0x1ae);
define('Type_EXTERNSHEET', 0x17);
define('Type_NAME', 0x18);
define('Type_MERGEDCELLS', 0xe5);
define('Type_SELECTION', 0x1d);
define('Type_FONT', 0x31);
define('Type_FORMAT', 0x041e);
define('Type_XF', 0xe0);
define('Type_DEFAULTROWHEIGHT', 0x225);
define('Type_FILEPASS', 0x2f);
define('Type_INTERFACEHDR', 0xe1);
define('Type_WRITEPROT', 0x86);
define('Type_FILESHARING', 0x5B);
define('Type_XCT', 0x59);
define('Type_CRN', 0x5a);

/**
* Class for regenerating Excel Spreadsheets
* @package Excel_Reviser
* @author kishiyan <excelreviser@gmail.com>
* @copyright Copyright &copy; 2006-2007, kishiyan
* @since PHP 4.4
* @example ./sample.php sample
*/
class Excel_Reviser
{
	// temp for workbook Globals Substream data
	var $wbdat='';
	// part of workbook Globals Substream data
	var $globaldat=array();
	// original rowrecord
	var $rowblock=array();
	// original colinfo record
	var $colblock=array();
	// buffer for all cell-record
	var $cellblock=array();
	// sheet-block data
	var $sheetbin=array();
	// each parameter of sheet record
	var $boundsheets=array();
	// buffer of user setting parameter
	var $revise_dat=array();
	// each data of shared string
	var $eachsst=array();
	// hyperlink-data by user
	var $hlink = array();
	// sheet-number for erase by user
	var $rmsheets=array();
	// cell for erase by user
	var $rmcells=array();
	// option for some type of formula
	var $exp_mode=0;
	// duplicate-sheet data by user
	var $dupsheet=array();
	// part of original sheet data
	var $stable=array();
	// charactor-set name
	var $charset;
	// option for graph-object reference
	var $opt_ref3d;
	// option for parse mode
	var $opt_parsemode;
	// mergecells data
	var $mergecells=array();
	// set/unset merge info
	var $mergeinfo=array();
	// set column width
	var $colwidth=array();
	// set row height
	var $rowheight=array();
	// font data
	var $recFONT=array();
	// format data
	var $recFORMAT=array();
	// XF data
	var $recXF=array();
	// DEFAULTROWHEIGHT data
	var $defrowH=array();
	// DEFCOLWIDTH data
	var $defcolW=array();
	// print Area data
	var $prnarea=array();
	// print Title data
	var $prntitle=array();
	// debug for image-data
	var $debug_image=1;
	// save magic_quotes Flag
	var $Flag_Magic_Quotes=False;
	// Error-handling method
	var $Flag_Error_Handling=0;
	// Error-Reporting Level
	var $Flag_Error_Reporting= E_ALL;
	// property inherit
	var $Flag_inherit_Info= 0;
	// Streams in OLE-container
	var $orgStreams=array();
	// File-Protection mode
	var $Mode_WR_Protect = 0;
	// Hash Value for template-read
	var $PW_Hash_read = NULL;
	// Hash Value for output-file
	var $PW_Hash_write = NULL;
	// Hash Value for output-file
	var $Flag_Read_Only = 0;
	// File Open Password for read
	var $openFPass = NULL;
	// File Open Password for write
	var $saveFPass = NULL;
	// command-path for de/crypto
	var $MkHashCmd = NULL;

	// Constructor
	function Excel_Reviser(){
//error_reporting(E_ALL ^ E_NOTICE);
		$this->charset = Default_CHARSET;
		$this->opt_ref3d = 0;
		$this->globaldat['presheet']='';
		$this->globaldat['presst']='';
		$this->globaldat['last']='';
		$this->globaldat['presup']='';
		$this->globaldat['supbook']='';
		$this->globaldat['extsheet']='';
		$this->globaldat['name']='';
		$this->globaldat['namerecord']='';
		$this->globaldat['exsstbin']='';
	}

	/**
	* Set(Get) internal charset, if you use multibyte-code.
	* @param string $chrset charactor-set name(Ex. SJIS)
	* @return string current charector-set name
	* @access public
	*/
	function setInternalCharset($chrset=''){
		if (strlen(trim($chrset)) > 2) {
			$this->charset = $chrset;
		}
		return $this->charset;
	}

	/**
	* Set(Get) parse mode, 1: include cell-attribute.
	* @param string $mode set parse mode
	* @return string current parse mode
	* @access public
	* @example ./sample_ex1.php sample_ex1
	*/
	function setParseMode($mode=null){
		if ($mode == 1) {
			$this->opt_parsemode = 1;
		}
		return $this->opt_parsemode;
	}

	/**
	* Set reference option for graph-object in added sheet
	* (This is experimental function. It operates under only
	*  the conditions which are limited.) 
	* @param integer $opt  1 = change link to self-sheet
	*                      0 = keep original link-address
	* @return integer current value
	* @access public
	* @example ./sample2.php sample2
	*/
	function setOptionRef3d($opt=null){
		if ($opt !== null){
			$this->opt_ref3d= $opt;
		}
		return $this->opt_ref3d;
	}

	/**
	* Set Cells to merge
	* @param integer $sn sheet-number  0 base indexed
	* @param integer $rowst row number for top-cell
	* @param integer $rowen row number for bottom-cell
	* @param integer $colst column number for left-cell
	* @param integer $colen column number for right-cell
	* @access public
	* @example ./sample3.php sample3
	*/
	function setCellMerge($sn,$rowst,$rowen,$colst,$colen){
		if ($sn < 0) return -1;
		if ($rowst < 0 || $rowst > 65535) return -1;
		if ($rowen < 0 || $rowen > 65535) return -1;
		if ($colst < 0 || $colst > 255) return -1;
		if ($colen < 0 || $colen > 255) return -1;
		if ($rowst == $rowen && $colst == $colen) return -1;
		if ($rowst > $rowen) return -1;
		if ($colst > $colen) return -1;
		$mtmp['rows']=$rowst;
		$mtmp['rowe']=$rowen;
		$mtmp['cols']=$colst;
		$mtmp['cole']=$colen;
		$this->mergeinfo['set'][$sn][]=$mtmp;
	}


	/**
	* Unset original MergedCells
	* @param integer  $sn sheet-number  0 base indexed
	* @access public
	* @example ./sample3.php sample3
	*/
	function unsetCellMerge($sn){
		if ($sn < 0) return -1;
		$this->mergeinfo['unset'][$sn]=TRUE;
	}


	/**
	* make MergedCells-info
	* @param integer  $sn sheet-number  0 base indexed
	* @access private
	*/
	function makeMergeinfo($sn){
		if ($sn < 0) return -1;
		if (isset($this->mergeinfo['unset'][$sn])) unset($this->mergecells[$sn]);
		if (isset($this->mergeinfo['set'][$sn])){
			foreach($this->mergeinfo['set'][$sn] as $val){
	            if (count($this->mergecells[$sn])) 
				foreach($this->mergecells[$sn] as $key=>$val0){
					if ($val['rows']==$val0['rows'] && $val['cols']==$val0['cols'])
						unset($this->mergecells[$sn][$key]);
				}
				$this->mergecells[$sn][]=$val;
			}
		}
	}


	/**
	* create file
	* @param string $outfile filename for web output
	* @param string $path if not null then save file
	* @access public
	* @example ./sample.php sample
	*/
	function createFile($outfile,$path=null){
		$res = $this->reviseFile('',$outfile,$path);
		if ($this->isError($res)) return $res;
	}

	/**
	* Parse file and Remake
	* @param string $readfile full path filename for read
	* @param string $outfile filename for web output
	* @param string $path if not null then save file
	* @access public
	* @example ./sample.php sample
	*/
	function reviseFile($readfile,$outfile,$path=null){
		$Flag_Error_Reporting = error_reporting();
		//error_reporting(E_ALL ^ E_NOTICE);
		error_reporting(E_ALL ^ (E_NOTICE | E_DEPRECATED));
		$res = $this->parseFile($readfile);
		if ($this->isError($res)) {
			error_reporting($Flag_Error_Reporting);
			return $res;
		}
		$res=$this->reviseCell();
		if ($this->isError($res)) return $res;
		$res = $this->makeFile($outfile,$path);
		error_reporting($Flag_Error_Reporting);
		if ($this->isError($res)) return $res;
	}

	/**
	* Remake file
	* @param string $outfile filename for web output
	* @param string $path if not null then save file
	* @access public
	* @example ./sample4.php sample4
	*/
	function buildFile($outfile,$path=null){
		$res=$this->reviseCell();
		if ($this->isError($res)) return $res;
		$res = $this->makeFile($outfile,$path);
		if ($this->isError($res)) return $res;
	}

	/**
	* Copy Sheet
	* @param integer $orgsn original sheet-number  0 indexed
	* @param integer $num number of sheet to duplicate
	* @access public
	* @example ./sample2.php sample2
	*/
	function addSheet($orgsn,$num){
		if ($num < 1) return;
		$tmp['orgsn']=$orgsn;
		$tmp['count']=$num;
		$this->dupsheet[]=$tmp;
	}

	/**
	* Add hyperlink to Cell
	* @param integer $sn sheet number
	* @param integer $row Row position
	* @param integer $col Column posion  0indexed
	* @param string $desc cell description(option)
	* @param string $link absolute link path
	* @param integer $refrow reference row(option)
	* @param integer $refcol reference column(option)
	* @param integer $refsheet reference sheet number(option)
	* @access public
	* @example ./sample.php sample
	*/
	function addHLink($sn,$row,$col,$desc='',$link, $refrow = null, $refcol = null, $refsheet = null){
		if (trim($link)=='') return;
		if ($desc == '') {
			$opt=0x03;
			$disp='';
			$desc=$link;
		} else {
			$opt=0x17;
			$str=mb_convert_encoding($desc,'UTF-16LE',$this->charset);
			$disp=pack("V",mb_strlen($str,'UTF-16LE')+1).$str."\x00\x00";
		}
		$link=mb_convert_encoding($link,'UTF-16LE',$this->charset)."\x00\x00";
		$linkrcd=pack("vvvv",$row,$row,$col,$col)
				. pack("H*","d0c9ea79f9bace118c8200aa004ba90b02000000")
				. pack("V",$opt).$disp
				. pack("H*","e0c9ea79f9bace118c8200aa004ba90b")
				. pack("V",strlen($link)).$link;
		$this->hlink[$sn][]=pack("vv",0x1b8,strlen($linkrcd)) . $linkrcd;
		$this->addString($sn,$row,$col,$desc, $refrow, $refcol, $refsheet);
	}

	/**
	* Set remove Sheet number
	* @param integer $sheet sheet number  0 indexed
	* @access public
	* @example ./sample.php sample
	*/
	function rmSheet($sheet){
		if (is_numeric($sheet)){
			$this->rmsheets[$sheet]=TRUE;
		}
	}

	/**
	* Set remove Cell
	* @param integer $sheet sheet number
	* @param integer $row Row position
	* @param integer $col Column posion  0 base indexed
	* @access public
	* @example ./sample.php sample
	*/
	function rmCell($sheet,$row,$col){
		if (is_numeric($sheet) && is_numeric($row) && is_numeric($col)){
			$this->rmcells[$sheet][$row][$col]=TRUE;
		}
	}

	/**
	* Set Row height
	* @param integer $sheet sheet number
	* @param integer $row Row position
	* @param integer $height Height of the row, in twips = 1/20 of a point
	* @since Ver0.21
	* @access public
	* @example ./sample3.php sample3
	*/
	function chgRowHeight($sheet,$row,$height){
		if (is_numeric($sheet) && is_numeric($row) && is_numeric($height)){
		if ($sheet < 0) return -1;
		if ($row < 0 || $row > 65535) return -1;
		if ($height <= 0) return -1;
			$this->rowheight[$sheet][$row]=$height;
		}
	}

	/**
	* Set Column width
	* @param integer $sheet sheet number
	* @param integer $col Column position
	* @param integer $width Width of the columns in 1/256 of the width of the zero character
	* @since Ver0.21
	* @access public
	* @example ./sample3.php sample3
	*/
	function chgColWidth($sheet,$col,$width){
		if (is_numeric($sheet) && is_numeric($col) && is_numeric($width)){
		if ($sheet < 0) return -1;
		if ($col < 0 || $col > 255) return -1;
		if ($width <= 0) return -1;
		$this->colwidth[$sheet][$col]=$width;
		}
	}

	/**
	* Add String to Cell
	* @param integer $sheet sheet number
	* @param integer $row Row position
	* @param integer $col Column posion  0indexed
	* @param string $str string
	* @param integer $refrow reference row(option)
	* @param integer $refcol reference column(option)
	* @param integer $refsheet reference sheet number(option)
	* @access public
	* @example ./sample.php sample
	*/
	function addString($sheet,$row, $col, $str, $refrow = null, $refcol = null, $refsheet = null){
		$val['sheet']=$sheet;
		$val['row']=$row;
		$val['col']=$col;
		$val['str']=$str;
		$val['refrow']=$refrow;
		$val['refcol']=$refcol;
		$val['refsheet']=$refsheet;
		$this->revise_dat['add_str'][]=$val;
	}

	/**
	* Add Number to Cell
	* @param integer $sheet sheet number
	* @param integer $row Row position
	* @param integer $col Column position  0indexed
	* @param integer $num number
	* @param integer $refrow reference row(option)
	* @param integer $refcol refernce column(option)
	* @param integer $refsheet reference sheet number(option)
	* @access public
	* @example ./sample.php sample
	*/
	function addNumber($sheet,$row, $col, $num, $refrow = null, $refcol = null, $refsheet = null){
		$val['sheet']=$sheet;
		$val['row']=$row;
		$val['col']=$col;
		$val['num']=$num;
		$val['refrow']=$refrow;
		$val['refcol']=$refcol;
		$val['refsheet']=$refsheet;
		$this->revise_dat['add_num'][]=$val;
	}

	/**
	* Add Formula to Cell
	* @param integer $sheet sheet number
	* @param integer $row Row position
	* @param integer $col Column position  0indexed
	* @param integer $formula Formula
	* @param integer $refrow reference row(option)
	* @param integer $refcol refernce column(option)
	* @param integer $refsheet reference sheet number(option)
	* @access public
	*/
	function addFormula($sheet,$row, $col, $formula, $refrow = null, $refcol = null, $refsheet = null, $opt=0){
		$val['sheet']=$sheet;
		$val['row']=$row;
		$val['col']=$col;
		$val['formula']=$formula;
		$val['refrow']=$refrow;
		$val['refcol']=$refcol;
		$val['refsheet']=$refsheet;
		$val['opt']=$opt;
		$this->revise_dat['add_formula'][]=$val;
	}

// TEST 2009.02.15
	/**
	* Copy Row-Line from Template
	* @param integer $sheetsrc sheet number for source
	* @param integer $rowsrc row position for source
	* @param integer $sheetdest sheet number for destination
	* @param integer $rowdest row position for destination
	* @param integer $num lot of lines for copy
	* @param integer $inc 0=source-line is fixed. 1=incriment for each line
	* @access public
	*/
	function copyRowLine($sheetsrc,$rowsrc,$sheetdest,$rowdest,$num,$inc=0){
		$val['sheetsrc']=$sheetsrc;
		$val['rowsrc']=$rowsrc;
		$val['sheetdest']=$sheetdest;
		$val['rowdest']=$rowdest;
		$val['num']=$num;
		$val['inc']=$inc;
		$this->revise_dat['copy_row'][]=$val;
	}

	/**
	* Copy Column-Line from Template
	* @param integer $sheetsrc sheet number for source
	* @param integer $colsrc column start position for source
	* @param integer $sheetdest sheet number for destination
	* @param integer $coldest column start position for destination
	* @param integer $num lot of lines for copy
	* @param integer $inc 0=source line is fixed. 1=incriment for each line
	* @access public
	*/
	function copyColLine($sheetsrc,$colsrc,$sheetdest,$coldest,$num,$inc=0){
		$val['sheetsrc']=$sheetsrc;
		$val['colsrc']=$colsrc;
		$val['sheetdest']=$sheetdest;
		$val['coldest']=$coldest;
		$val['num']=$num;
		$val['inc']=$inc;
		$this->revise_dat['copy_col'][]=$val;
	}
// TEST end

	/**
	* Add Formula-Record to Cell by Direct
	*  This is Experimental method  2007/10/13
	*
	* @param integer $sheet sheet number
	* @param integer $row Row position
	* @param integer $col Column position  0 indexed
	* @param integer $record (formula binary-record)
	* @param integer $refrow reference row(option)
	* @param integer $refcol refernce column(option)
	* @param integer $refsheet reference sheet number(option)
	* @access private
	*/
	function addFormulaRecord($sheet,$row, $col, $record, $refrow = null, $refcol = null, $refsheet = null){
		if (mb_check_encoding($record,'ASCII') && !preg_match('/[\x00-\x1f]/', $record)) {
			$this->addString($sheet,$row, $col, $record, $refrow, $refcol, $refsheet);
			return;
		}
		if (strlen($record) < 1) return -1;
		if ($sheet < 0 || $row < 0 || $col < 0 || $refsheet < 0 || $refrow < 0 || $refcol < 0) return -1;
		$formlen = strlen($record);
		$header	= pack("vv", 0x06, 0x16 + $formlen);
		$data	  = pack("vvvdvVv", $row, $col, $this->_getcolxf($sheet,$col), 0, 3, 9, $formlen);
		$val['sheet']=$sheet;
		$val['row']=$row;
		$val['col']=$col;
		$val['record']=$header.$data.$record;
		$val['refrow']=$refrow;
		$val['refcol']=$refcol;
		$val['refsheet']=$refsheet;
		$this->revise_dat['option'][]=$val;
	}

	/**
	* Change original string to new string
	* @param string $org original String
	* @param string $new new string
	* @access public
	* @example ./sample.php sample
	*/
	function changeStr($org, $new){
		if ($new == '') $new = ' ';
		if (mb_detect_encoding($org,"ASCII,".$this->charset.",ISO-8859-1") == 'ISO-8859-1')
			$org=mb_convert_encoding($org,$this->charset,'auto');
		if (mb_detect_encoding($new,"ASCII,".$this->charset.",ISO-8859-1") == 'ISO-8859-1')
			$new=mb_convert_encoding($new,$this->charset,'auto');
		$tmp['org']=mb_convert_encoding($org,'UTF-16LE',$this->charset);
		$tmp['new']=mb_convert_encoding($new,'UTF-16LE',$this->charset);
		$this->revise_dat['replace'][]=$tmp;
	}

	/**
	* overwrite Sheetname
	* @param integer $sn sheet number
	* @param string $str new sheet name
	* @access public
	* @example ./sample.php sample
	*/
	function setSheetname($sn,$str){
			$len = strlen($str);
			if (mb_detect_encoding($str,"ASCII,ISO-8859-1")=="ASCII"){
				$opt =0;
			} else {
				$opt =1;
				$str = mb_convert_encoding($str,'UTF-16LE',$this->charset);
				$len = mb_strlen($str,'UTF-16LE');
			}
			$val = pack("CC",$len,$opt);
		$this->revise_dat['sheetname'][$sn]=$val.$str;
	}

	/**
	* overwrite header string
	* @param integer $sn sheet number
	* @param string $str new header-string
	* @access public
	* @example ./sample.php sample
	*/
	function setHeader($sn,$str){
			if (mb_detect_encoding($str,"ASCII,ISO-8859-1")=="ASCII"){
				$opt =0;
				$len = strlen($str);
			} else {
				$opt =1;
				$str = mb_convert_encoding($str,'UTF-16LE',$this->charset);
				$len = mb_strlen($str,'UTF-16LE');
			}
			$val = pack("vC",$len,$opt);
		$this->revise_dat['header'][$sn]=$val.$str;
	}

	/**
	* overwrite footer string
	* @param integer $sn sheet number
	* @param string $str new footer-string
	* @access public
	* @example ./sample.php sample
	*/
	function setFooter($sn,$str){
			if (mb_detect_encoding($str,"ASCII,ISO-8859-1")=="ASCII"){
				$opt =0;
				$len = strlen($str);
			} else {
				$opt =1;
				$str = mb_convert_encoding($str,'UTF-16LE',$this->charset);
				$len = mb_strlen($str,'UTF-16LE');
			}
			$val = pack("vC",$len,$opt);
		$this->revise_dat['footer'][$sn]=$val.$str;
	}

	/**
	* Add Blank Cell
	* @param integer $sheet sheet number  0 base indexed
	* @param integer $row Row position  0 base indexed
	* @param integer $col Column posion  0 base indexed
	* @param integer $refrow reference row(option)
	* @param integer $refcol reference column(option)
	* @param integer $refsheet ref sheet number(option)
	* @access public
	* @example ./sample3.php sample3
	*/
	function addBlank($sheet,$row, $col, $refrow, $refcol, $refsheet = null){
		$val['sheet']=$sheet;
		$val['row']=$row;
		$val['col']=$col;
		$val['refrow']=$refrow;
		$val['refcol']=$refcol;
		$val['refsheet']=$refsheet;
		$this->revise_dat['add_blank'][]=$val;
	}

	/**
	* Set Printtitle
	* @param integer $sheet sheet number  0 base indexed
	* @param integer $row1st First Row position  0 base indexed
	* @param integer $rowlast Last Row position  0 base indexed
	* @param integer $col1st First Column position  0 base indexed
	* @param integer $collast Last Column position  0 base indexed
	* @access public
	* @example ./sample4.php sample4
	*/
	function setPrintTitle($sheet,$row1st=null,$rowlast=null,$col1st=null,$collast=null){
		if ($sheet < 0) return;
		if ($row1st===null && $col1st===null) return;
		if ($rowlast===null) $rowlast=$row1st;
		if ($collast===null) $collast=$col1st;
		if ($row1st!==null) if ($row1st > $rowlast) return;
		if ($col1st!==null) if ($col1st > $collast) return;
		$this->prntitle[$sheet]['row1st']=$row1st;
		$this->prntitle[$sheet]['rowlast']=$rowlast;
		$this->prntitle[$sheet]['col1st']=$col1st;
		$this->prntitle[$sheet]['collast']=$collast;
	}

	/**
	* Set Row Print-title
	* @param integer $sheet sheet number  0 base indexed
	* @param integer $row1st First Row position  0 base indexed
	* @param integer $rowlast Last Row position  0 base indexed
	* @access public
	* @example ./sample4.php sample4
	*/
	function setPrintTitleRow($sheet,$row1st,$rowlast=null){
		if ($sheet < 0) return;
		if ($row1st===null) return;
		if ($rowlast===null) $rowlast=$row1st;
		if ($row1st > $rowlast) return;
		$this->prntitle[$sheet]['row1st']=$row1st;
		$this->prntitle[$sheet]['rowlast']=$rowlast;
	}

	/**
	* Set Column Printtitle
	* @param integer $sheet sheet number  0 base indexed
	* @param integer $col1st First Column position  0 base indexed
	* @param integer $collast Last Column position  0 base indexed
	* @access public
	* @example ./sample4.php sample4
	*/
	function setPrintTitleCol($sheet,$col1st,$collast=null){
		if ($sheet < 0) return;
		if ($col1st===null) return;
		if ($collast===null) $collast=$col1st;
		if ($col1st > $collast) return;
		$this->prntitle[$sheet]['col1st']=$col1st;
		$this->prntitle[$sheet]['collast']=$collast;
	}

	/**
	* Set PrintArea
	* @param integer $sheet sheet number  0 base indexed
	* @param integer $row1st First Row position  0 base indexed
	* @param integer $rowlast Last Row position  0 base indexed
	* @param integer $col1st First Column position  0 base indexed
	* @param integer $collast Last Column position  0 base indexed
	* @access public
	* @example ./sample4.php sample4
	*/
	function setPrintArea($sheet,$row1st,$rowlast,$col1st,$collast){
		if ($sheet < 0) return;
		if ($row1st>$rowlast || $col1st>$collast) return;
		$this->prnarea[$sheet]['row1st']=$row1st;
		$this->prnarea[$sheet]['rowlast']=$rowlast;
		$this->prnarea[$sheet]['col1st']=$col1st;
		$this->prnarea[$sheet]['collast']=$collast;
	}

	/**
	* Set Password for a read-only file.
	* @param string $pass password for write-protected template file
	* @param integer $readonly 1: Recommend read-only state while loading the file
	* @access public
	*/
	function setWritePassword($pass){
		if (strlen($pass) < 1) return;
		if (strlen($pass) > 16) return;
		$this->PW_Hash_read=$this->makehash16b($pass);
	}

	/**
	* Set(Unset) read-only password
	* @param string $pass password for read-only file to output
    *                    if NULL then force unprotect
	* @access public
	*/
	function setWriteProtect($pass=null, $readonly=0){
		if ($readonly==1) $this->Flag_Read_Only=1;
		if ($pass===null || $pass=="") {
			$this->PW_Hash_write=-1;
			return;
		}
		if (strlen($pass) < 1) return;
		if (strlen($pass) > 15) return;
		$this->PW_Hash_write=$this->makehash16b($pass);
	}

	/**
	* Set File Open password
	* @param string $pass file-open password for template
	* @access public
	*/
	function setOpenFilePassword($pass=''){
		if (strlen(trim($pass)) > 0) {
			$this->openFPass = $pass;
		}
		return $this->openFPass;
	}

	/**
	* Set File Open password
	* @param string $pass file-open password for output-file
	* @access public
	*/
	function setSaveFilePassword($pass=''){
		if (strlen(trim($pass)) > 0) {
			$this->saveFPass = $pass;
		}
		return $this->saveFPass;
	}


	/**
	* Add Blank Cell
	* @param integer $sheet sheet number
	* @param integer $row Row position
	* @param integer $col Column posion  0indexed
	* @param integer $refrow reference row(option)
	* @param integer $refcol reference column(option)
	* @param integer $refsheet reference sheet number(option)
	* @access private
	*/
	function _addBlank($sheet,$row, $col, $refrow, $refcol, $refsheet = null){
		if (($row < 0) || ($col < 0) || ($sheet < 0)) return -1;
		if (($refrow < 0) || ($refcol < 0)) return -1;
		if ($refsheet === null) $refsheet = $sheet;
		$xf= (isset($this->cellblock[$refsheet][$refrow][$refcol]['xf'])) ? $this->cellblock[$refsheet][$refrow][$refcol]['xf'] : $this->_getcolxf($refsheet,$refcol);
		$header    = pack('vv', Type_BLANK, 0x06);
		$data      = pack('vvv', $row, $col, $xf);
		$this->cellblock[$sheet][$row][$col]['xf']=$xf;
		$this->cellblock[$sheet][$row][$col]['type']=Type_BLANK;
		$this->cellblock[$sheet][$row][$col]['dat']='';
		$this->cellblock[$sheet][$row][$col]['record']=bin2hex($header.$data);
	}

	/**
	* Add String to Cell (for internal access)
	* @param  $sn:sheet number,$row:Row position,$col:Column posion  0indexed
	* @param  $str:string
	* @param  $refrow:referrence row(option)
	* @param  $refcol:ref column(option)
	* @param  $refsheet:ref sheet number(option)
	* @access private
	*/
	function _addString($sheet,$row, $col, $str, $refrow = null, $refcol = null, $refsheet = null){
		if (($row < 0) || ($col < 0) || ($sheet < 0)) return -1;
		if ($refsheet === null) $refsheet = $sheet;
		if (($refrow !== null) && ($refcol !== null)) {
			$xf= (isset($this->cellblock[$refsheet][$refrow][$refcol]['xf'])) ? $this->cellblock[$refsheet][$refrow][$refcol]['xf'] : $this->_getcolxf($refsheet,$refcol);
		} else {
			$xf= (isset($this->cellblock[$sheet][$row][$col]['xf'])) ? $this->cellblock[$sheet][$row][$col]['xf'] : $this->_getcolxf($sheet,$col);
		}
		if (mb_detect_encoding($str,"ASCII,ISO-8859-1")=="ASCII"){
			$opt =0;
			$str = mb_convert_encoding($str, "UTF-16LE", "ASCII");
		} else {
			$opt =1;
			$str = mb_convert_encoding($str, "UTF-16LE", $this->charset);
		}
		$len = mb_strlen($str, 'UTF-16LE');
		$tempsst['len']=$len;
		$tempsst['opt']=$opt;
		$tempsst['rtn']=0;
		$tempsst['apn']=0;
		$tempsst['str']=bin2hex($str);
		$tempsst['rt']='';
		$tempsst['ap']='';
		$this->eachsst[]=$tempsst;
		$header    = pack('vv', Type_LABELSST, 0x0a);
		$data      = pack('vvvV', $row, $col, $xf, count($this->eachsst)-1);
		$this->cellblock[$sheet][$row][$col]['xf']=$xf;
		$this->cellblock[$sheet][$row][$col]['type']=Type_LABELSST;
		$this->cellblock[$sheet][$row][$col]['dat']=bin2hex(pack("V",count($this->eachsst)-1));
		$this->cellblock[$sheet][$row][$col]['record']=bin2hex($header.$data);
		return;
	}

	/**
	* Add Number to Cell
	* @param  $sn:sheet number,$row:Row position,$col:column posion  0indexed
	* @param  $num:number
	* @param  $refrow:referrence row(option), $refcol:ref column(option)
	* @param  $refsheet:ref sheet number(option)
	* @access private
	*/
	function _addNumber($sheet,$row, $col, $num, $refrow = null, $refcol = null, $refsheet = null){
		if (($row < 0) || ($col < 0) || ($sheet < 0)) return -1;
		if ($refsheet === null) $refsheet = $sheet;
		if (($refrow !== null) && ($refcol !== null)) {
			$xf= (isset($this->cellblock[$refsheet][$refrow][$refcol]['xf'])) ? $this->cellblock[$refsheet][$refrow][$refcol]['xf'] : $this->_getcolxf($refsheet,$refcol);
		} else {
			$xf= (isset($this->cellblock[$sheet][$row][$col]['xf'])) ? $this->cellblock[$sheet][$row][$col]['xf'] : $this->_getcolxf($sheet,$col);
		}
		$packednum = (pack("N",1)==pack("L",1)) ? strrev(pack("d", $num)) : pack("d", $num); // added 
		$header    = pack('vv', Type_NUMBER, 0x0e);
//		$data      = pack('vvvd', $row, $col, $xf, $num);
	$data      = pack('vvv', $row, $col, $xf).$packednum; // edited 
		$this->cellblock[$sheet][$row][$col]['xf']=$xf;
		$this->cellblock[$sheet][$row][$col]['type']=Type_NUMBER;
//		$this->cellblock[$sheet][$row][$col]['dat']=bin2hex(pack("d", $num));
	$this->cellblock[$sheet][$row][$col]['dat']=bin2hex($packednum); //edited 
		$this->cellblock[$sheet][$row][$col]['record']=bin2hex($header.$data);
		return;
	}

	/**
	* Add Formula to Cell
	* @access private
	*/
	function _addFormula($sheet,$row, $col, $formula, $refrow = null, $refcol = null, $refsheet = null, $opt=0){
		if (($row < 0) || ($col < 0) || ($sheet < 0)) return -1;
		if ($refsheet === null) $refsheet = $sheet;
		if (($refrow !== null) && ($refcol !== null)) {
			$xf= (isset($this->cellblock[$refsheet][$refrow][$refcol]['xf'])) ? $this->cellblock[$refsheet][$refrow][$refcol]['xf'] : $this->_getcolxf($refsheet,$refcol);
		} else {
			$xf= (isset($this->cellblock[$sheet][$row][$col]['xf'])) ? $this->cellblock[$sheet][$row][$col]['xf'] : $this->_getcolxf($sheet,$col);
		}
// TODO make formura record

	$mkform = & new Formula_Parser();
	$formula = preg_replace("/(^[=@])/","",$formula);
	$error = $mkform->parse($formula);
	if ($mkform->isError($error)) {
		if ($opt==1){
			$errmes=$error->getMessage();
			$this->_addString($sheet,$row, $col, $errmes, $refrow, $refcol, $refsheet);
			return;
		} else {
			return $this->raiseError("ERR addFormula : ".$error->getMessage());
		}
	}
//	$record = $mkform->toReversePolish();
	$record = $mkform->convFormRecord();
	if ($mkform->isError($record)) {
		if ($opt==1){
			$this->_addString($sheet,$row, $col, $record->getMessage(), $refrow, $refcol, $refsheet);
		} else {
			return $this->raiseError("ERR addFormula : ".$record->getMessage());
		}
	}
		
		$formlen = strlen($record);
		$header    = pack('vv', Type_FORMULA2, 0x16 + $formlen);
		$data      = pack('vvvdvVv', $row, $col, $xf, 0, 3, 9, $formlen).$record;
		$this->cellblock[$sheet][$row][$col]['xf']=$xf;
		$this->cellblock[$sheet][$row][$col]['type']=Type_FORMULA2;
		$this->cellblock[$sheet][$row][$col]['dat']=bin2hex($data);
		$this->cellblock[$sheet][$row][$col]['record']=bin2hex($header.$data);
		return;
	}

	/**
	* read OLE container
	* @param  $Fname:filename
	* @access private
	*/
	function __oleread($Fname){
		if(!is_readable($Fname)) {
			return $this->raiseError("ERROR Cannot read file ${Fname} \nProbably there is not reading permission whether there is not a file");
		}
	// 2007.11.19
		$this->Flag_Magic_Quotes = get_magic_quotes_runtime();
		if ($this->Flag_Magic_Quotes) set_magic_quotes_runtime(0);
		$ole_data = @file_get_contents($Fname);
		if ($this->Flag_Magic_Quotes) set_magic_quotes_runtime($this->Flag_Magic_Quotes);
		if (!$ole_data) { 
			return $this->raiseError("ERROR Cannot open file ${Fname} \n");
		}
		if (substr($ole_data, 0, 8) != pack("CCCCCCCC",0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1)) {
			return $this->raiseError("ERROR Template file(${Fname}) is not EXCEL file.\n");
	   	}
		$numDepots = $this->__get4($ole_data, 0x2c);
		$sStartBlk = $this->__get4($ole_data, 0x3c);
		$ExBlock = $this->__get4($ole_data, 0x44);
		$numExBlks = $this->__get4($ole_data, 0x48);

		$len_ole = strlen($ole_data);
		if ($numDepots > ($len_ole / 65536 +1))
			return $this->raiseError("ERROR file($Fname) is broken (numDepots:${numDepots})");
		if ($sStartBlk > ($len_ole / 512 +1))
			return $this->raiseError("ERROR file($Fname) is broken (sStartBlk:${sStartBlk})");
		if ($ExBlock > ($len_ole / 512 +1))
			return $this->raiseError("ERROR file($Fname) is broken (ExBlock:${ExBlock})");
		if ($numExBlks > ($len_ole / 512 +1))
			return $this->raiseError("ERROR file($Fname) is broken (numExBlks:${numExBlks})");

		$DepotBlks = array();
		$pos = 0x4c;
		$dBlks = $numDepots;
		if ($numExBlks != 0) $dBlks = (0x200 - 0x4c)/4;
		for ($i = 0; $i < $dBlks; $i++) {
			$DepotBlks[$i] = $this->__get4($ole_data, $pos);
			$pos += 4;
		}

		for ($j = 0; $j < $numExBlks; $j++) {
			$pos = ($ExBlock + 1) * 0x200;
			$ReadBlks = min($numDepots - $dBlks, 0x200 / 4 - 1);
			for ($i = $dBlks; $i < $dBlks + $ReadBlks; $i++) {
				$DepotBlks[$i] = $this->__get4($ole_data, $pos);
				$pos += 4;
			}   
			$dBlks += $ReadBlks;
			if ($dBlks < $numDepots) $ExBlock = $this->__get4($ole_data, $pos);
		}

		$pos = 0;
		$index = 0;
		$BlkChain = array();
		for ($i = 0; $i < $numDepots; $i++) {
			$pos = ($DepotBlks[$i] + 1) * 0x200;
			for ($j = 0 ; $j < 0x200 / 4; $j++) {
				$BlkChain[$index] = $this->__get4($ole_data, $pos);
				$pos += 4 ;
				$index++;
			}
		}
		$eoc=pack("H*","FEFFFFFF");
		$eoc= $this->__get4($eoc,0);
		$pos = 0;
		$index = 0;
		$sBlkChain = array();
		while ($sStartBlk != $eoc) {
			$pos = ($sStartBlk + 1) * 0x200;
			for ($j = 0; $j < 0x80; $j++) {
				$sBlkChain[$index] = $this->__get4($ole_data, $pos);
				$pos += 4 ;
				$index++;
			}
			$chk[$sStartBlk]=true;
			$sStartBlk = $BlkChain[$sStartBlk];
			if(isset($chk[$sStartBlk])){
	return $this->raiseError("Big Block chain for small-block ERROR 1\nTemplate file is broken");
			}
		}
		unset($chk);
		$block = $this->__get4($ole_data, 0x30);
		$pos = 0;
		$entry = '';
		while ($block != $eoc)  {
			$pos = ($block + 1) * 0x200;
			$entry .= substr($ole_data, $pos, 0x200);
			$chk[$block]=true;
			$block = $BlkChain[$block];
			if(isset($chk[$block])){
	return $this->raiseError("Big Block chain for Entry  ERROR 2\nTemplate file is broken");
			}
		}
		unset($chk);
		$offset = 0;
		$bookKey=0;
		$tmpDir=array();
		$rootBlock =$this->__get4($entry, 0x74);
		while ($offset < strlen($entry)) {
			  $d = substr($entry, $offset, 0x80);
			  $name = str_replace("\x00", "", substr($d,0,$this->__get2($d,0x40)));
			if (($name == "Workbook") || ($name == "Book")) {
				$wbstartBlock =$this->__get4($d, 0x74);
				$wbsize = $this->__get4($d, 0x78);
			}
			if ($name == "Root Entry" || $name == "R") {
//				$rootBlock =$this->__get4($d, 0x74);
			} else if (strlen($name)>0){
				$tmpDir['startB']=$this->__get4($d, 0x74);
				$tmpDir['size']=$this->__get4($d, 0x78);
				$tmpDir['dat']='';
				if (($name == "Workbook") || ($name == "Book")) $bookKey=$name;
				if ($this->Flag_inherit_Info != 1){
					if (($name == "Workbook") || ($name == "Book")) $this->orgStreams[$name]=$tmpDir;
				} else {
					$this->orgStreams[$name]=$tmpDir;
				}
			}
			$offset += 0x80;
		}

		if (! isset($rootBlock)) return $this->raiseError("Unknown OLE-type. Can't find Root-Entry");
		$pos = 0;
		$rdata = '';
		while ($rootBlock != $eoc)  {
			$pos = ($rootBlock + 1) * 0x200;
			$rdata .= substr($ole_data, $pos, 0x200);
			$chk[$rootBlock]=true;
			$rootBlock = $BlkChain[$rootBlock];
			if(isset($chk[$rootBlock])){
				return $this->raiseError("Root Block chain read ERROR 2.1\n  Template file is broken");
			}
			unset($chk);
		}
		foreach($this->orgStreams as $name=>$tdir){
			if ($tdir['size'] <1) continue;
			if ($tdir['size'] < 0x1000) {
				$pos = 0;
				$tData = '';
				$block = $tdir['startB'];
				while ($block != $eoc) {
					$pos = $block * 0x40;
					$tData .= substr($rdata, $pos, 0x40);
					$chk[$block]=true;
					$block = $sBlkChain[$block];
					if(isset($chk[$block])){
		return $this->raiseError("Root Block chain read ERROR 2.2\n  Template file is broken");
					}
				}
				unset($chk);
				$this->orgStreams[$name]['dat'] = $tData;
			} else {
				$numBlocks = ($tdir['size'] + 0x1ff) / 0x200;
				if ($numBlocks == 0) continue;
				$tData = '';
				$block = $tdir['startB'];
				$pos = 0;
				while ($block != $eoc) {
					$pos = ($block + 1) * 0x200;
					$tData .= substr($ole_data, $pos, 0x200);
					$chk[$block]=true;
					$block = $BlkChain[$block];
					if(isset($chk[$block])){
					return $this->raiseError("Big Block chain ERROR 3\nTemplate file is broken");
					}
				}
				unset($chk);
				$this->orgStreams[$name]['dat'] = $tData;
			}
		}
		return $this->orgStreams[$bookKey]['dat'];
	}

	/**
	* parse sheetblock
	* @access private
	*/
	function __parsesheet(&$dat,$sn,$spos){
		$code = 0;
		$version = $this->__get2($dat,$spos + 4);
		$substreamType = $this->__get2($dat,$spos + 6);
		if ($version != Code_BIFF8) {
			return $this->raiseError("Contents(included sheet) is not BIFF8 format.\n");
		}
		if ($substreamType != Code_Worksheet) {
			return $this->raiseError("Contents is unknown format.\nCan't find Worksheet.\n");
		}
		$tmp='';
		$dimnum=0;
		$bof_num=0;
		$sposlimit=strlen($dat);
		while($code != Type_EOF) {
			if ($spos > $sposlimit) {
				return $this->raiseError("Sheet $sn Read ERROR\nTemplate file is broken.\n");
			}
			$code = $this->__get2($dat,$spos);
			$length = $this->__get2($dat,$spos + 2);
			if ($code == Type_BOF) $bof_num++;
			if ($bof_num > 1){
				$tmp.=substr($dat, $spos, $length+4);
				while($code != Type_EOF) {
					if ($spos > $sposlimit) {
						return $this->raiseError("Parse-Sheet Error\n");
					}
					$spos += $length+4;
					$code = $this->__get2($dat,$spos);
					$length = $this->__get2($dat,$spos + 2);
					$tmp.=substr($dat, $spos, $length+4);
				}
				$bof_num--;
				$spos += $length+4;
				$code = $this->__get2($dat,$spos);
				$length = $this->__get2($dat,$spos + 2);
				$tmp.=substr($dat, $spos, $length+4);
			}else
			switch ($code) {
				case Type_HEADER:
					if ($tmp) {
						$this->sheetbin[$sn]['preHF']=$tmp;
						$tmp='';
					}
					$this->sheetbin[$sn]['header']=substr($dat, $spos, $length+4);
					break;
				case Type_FOOTER:
//					if ($tmp) {
//						$this->sheetbin[$sn]['preHF']=$tmp;
//						$tmp='';
//					}
					$this->sheetbin[$sn]['footer']=substr($dat, $spos, $length+4);
					break;
				case Type_DEFCOLWIDTH:
					$tmp.=substr($dat, $spos, $length+4);
					$this->sheetbin[$sn]['preBT']=$tmp;
					$tmp='';

					$this->defcolW[$sn]=$this->__get2($dat,$spos+4);
					break;
				case Type_DEFAULTROWHEIGHT:
					$tmp.=substr($dat, $spos, $length+4);
					$this->defrowH[$sn]=$this->__get2($dat,$spos+6);
					break;
				case Type_COLINFO:
					$work['head']=substr($dat, $spos, 4);
					$colst=$this->__get2($dat,$spos + 4);
					$colen=$this->__get2($dat,$spos + 6);
					if ($colen >255) $colen=255;
					$work['width']=$this->__get2($dat,$spos + 8);
					$work['xf']=$this->__get2($dat,$spos + 10);
					$work['opt']=$this->__get2($dat,$spos + 12);
					$work['unk']=$this->__get2($dat,$spos + 14);
					for ($i=$colst;$i<=$colen;$i++){
						$work['colst']=$i;
						$work['colen']=$i;
						$work['all']=substr($dat, $spos, 4);
						$work['all'].=pack("v",$i).pack("v",$i);
						$work['all'].=substr($dat, $spos+8, $length-4);
						$this->colblock[$sn][$i]=$work;
					}
					unset($work);
					break;
				case Type_DIMENSION:
					$tmp.=substr($dat, $spos, $length+4);
	if ($dimnum==0){
					$this->sheetbin[$sn]['preCB']=$tmp;
					$tmp='';
	}
	$dimnum++;
					break;
				case Type_ROW:
					$row=$this->__get2($dat,$spos + 4);
					$this->rowblock[$sn][$row]['rowhead']=bin2hex(substr($dat, $spos, 4));
					$this->rowblock[$sn][$row]['col1st']=$this->__get2($dat,$spos + 6);
					$this->rowblock[$sn][$row]['collast']=$this->__get2($dat,$spos + 8);
					$this->rowblock[$sn][$row]['height']=$this->__get2($dat,$spos + 10);
					$this->rowblock[$sn][$row]['notused0']=$this->__get2($dat,$spos + 12);
					$this->rowblock[$sn][$row]['notused1']=$this->__get2($dat,$spos + 14);
					$this->rowblock[$sn][$row]['opt0']=$this->__get2($dat,$spos + 16);
					$this->rowblock[$sn][$row]['opt1']=$this->__get2($dat,$spos + 18);
					break;
				case Type_RK2:
				case Type_LABEL:
				case Type_LABELSST:
				case Type_NUMBER:
				case Type_FORMULA2:
				case Type_BOOLERR:
				case Type_BLANK:
					$row=$this->__get2($dat,$spos + 4);
					$col=$this->__get2($dat,$spos + 6);
					$this->cellblock[$sn][$row][$col]['xf']=$this->__get2($dat,$spos + 8);
					$this->cellblock[$sn][$row][$col]['type']=$code;
					$this->cellblock[$sn][$row][$col]['dat']=bin2hex(substr($dat, $spos+10, $length-6));
					$this->cellblock[$sn][$row][$col]['record']=bin2hex(substr($dat, $spos, $length+4));
					$this->cellblock[$sn][$row][$col]['string']='';
	if ($code == Type_FORMULA2){
		$dispnum = substr($dat, $spos+10, 8);
		$opflag = $this->__get2($dat,$spos + 18) | 0x02; // Calculate on open
		$tokens = substr($dat, $spos+20, $length - 16);
		$this->cellblock[$sn][$row][$col]['dat']=bin2hex($dispnum . pack("v",$opflag) . $tokens);
		$this->cellblock[$sn][$row][$col]['record']='';
		if ($this->exp_mode & 0x01) {
			if ($this->__get2($dat,$spos + $length + 4) == Type_SharedFormula){
				$spos += $length + 4;
				$length = $this->__get2($dat,$spos + 2);
				$sharedform[$row][$col]['firstR'] = $this->__get2($dat,$spos + 4);
				$sharedform[$row][$col]['lastR'] = $this->__get2($dat,$spos + 6);
				$sharedform[$row][$col]['firstC'] = $this->__get1($dat,$spos + 8);
				$sharedform[$row][$col]['lastC'] = $this->__get1($dat,$spos + 9);
				$sharedform[$row][$col]['formula'] = bin2hex(substr($dat,$spos+ 12,$length-8));
				$cur[$row][$col]=$this->__detrelcel($this->cellblock[$sn][$row-1][$col]['dat'],$sharedform[$row][$col]['formula'],$row-1,$col);
			}
			$sfdat=pack("H*", $this->cellblock[$sn][$row][$col]['dat']);
			if((($this->__get2($sfdat,8) & 8) == 8) && ($this->__get2($sfdat,14) == 5) && ($this->__get1($sfdat,16) == 1)){
				$refr =$this->__get2($sfdat,17);
		        	$refc =$this->__get2($sfdat,19);
				if (isset($sharedform[$refr][$refc]['formula'])){
					$this->cellblock[$sn][$row][$col]['record']='';
					$this->cellblock[$sn][$row][$col]['dat']=substr($sfdat, 0, 8);
					$this->cellblock[$sn][$row][$col]['dat'].=pack('v',0);
					$this->cellblock[$sn][$row][$col]['dat'].=substr($sfdat, 10, 4);
					$this->cellblock[$sn][$row][$col]['dat']=bin2hex($this->cellblock[$sn][$row][$col]['dat']);
	//			$this->cellblock[$sn][$row][$col]['dat'].=$sharedform[$refr][$refc]['formula'];
					$this->cellblock[$sn][$row][$col]['dat'].=bin2hex($this->__editformula($cur[$refr][$refc], pack("H*",$sharedform[$refr][$refc]['formula']), $row, $col));
				}
			}
		} else {
			if ($this->__get2($dat,$spos + $length + 4) == Type_SharedFormula){
				$spos += $length + 4;
				$length = $this->__get2($dat,$spos + 2);
				$this->cellblock[$sn][$row][$col]['sharedform']=substr($dat,$spos,$length+4);
			}
	  	}
		if ($this->__get2($dat,$spos + $length + 4) == Type_STRING){
			$spos += $length + 4;
			$length = $this->__get2($dat,$spos + 2);
			$this->cellblock[$sn][$row][$col]['string']=substr($dat,$spos,$length+4);
		}
	}
	
					break;
				case Type_MULBLANK:
					$muln=($length-6)/2;
					$row=$this->__get2($dat,$spos + 4);
					$col=$this->__get2($dat,$spos + 6);
					$i=-1;
					while(++$i < $muln){
						$this->cellblock[$sn][$row][$i+$col]['xf']=$this->__get2($dat,$spos+8+$i*2);
						$this->cellblock[$sn][$row][$i+$col]['type']=Type_BLANK;
						$this->cellblock[$sn][$row][$i+$col]['dat']='';
						$this->cellblock[$sn][$row][$i+$col]['record']=bin2hex(pack("vvvv", 0x0201, 0x06, $row, $i+$col). substr($dat, $spos+8+$i*2, 2));
					}
					break;
				case Type_MULRK:
					$muln=($length-6)/6;
					$row=$this->__get2($dat,$spos + 4);
					$col=$this->__get2($dat,$spos + 6);
					$i=-1;
					while(++$i < $muln){
						$this->cellblock[$sn][$row][$i+$col]['xf']=$this->__get2($dat,$spos+8+$i*6);
						$this->cellblock[$sn][$row][$i+$col]['type']=Type_RK;
						$this->cellblock[$sn][$row][$i+$col]['dat']=bin2hex(substr($dat, $spos+10+$i*6, 4));
						$this->cellblock[$sn][$row][$i+$col]['record']=bin2hex(pack("vvvv", 0x027e, 0x0a, $row, $i+$col). substr($dat, $spos+8+$i*6, 6));
					}
					break;
				case Type_MERGEDCELLS:
					$numrange=$this->__get2($dat,$spos+4);
					for($i=0;$i<$numrange;$i++){
						$mtmp['rows']=$this->__get2($dat,$spos+6+$i*8);
						$mtmp['rowe']=$this->__get2($dat,$spos+8+$i*8);
						$mtmp['cols']=$this->__get2($dat,$spos+10+$i*8);
						$mtmp['cole']=$this->__get2($dat,$spos+12+$i*8);
						$this->mergecells[$sn][]=$mtmp;
					}
					break;
				case Type_SELECTION:
					$tmp.= substr($dat, $spos, $length+4);
					if ($this->__get2($dat,$spos+$length+4)==Type_SELECTION) break;
					$this->sheetbin[$sn]['preMG']=$tmp;
					$tmp='';
					break;
				case Type_DBCELL:
					break;
				case Type_BUTTONPROPERTYSET:
					break;
				case Type_EOF:
					break;
				default:
					$tmp.= substr($dat, $spos, $length+4);
			}
			$spos += $length+4;
		}
		$this->sheetbin[$sn]['tail']=$tmp;
	}

	/**
	* detect some type of relative token in formula-record
	* @access private
	*/
	function __detrelcel($org,$share,$row,$col){
		$org=pack("H*",$org);
		$org=substr($org,14);
		$share=pack("H*",$share);
		$lenorg=strlen($org);
		$lenshare=strlen($share);
		if ($lenorg != $lenshare) return;
		$i=0;
		while($i < $lenorg - 3){
			if (($this->__get1($org,$i)==0x44) && ($this->__get1($share,$i)==0x4c)){
				if (((($this->__get2($org,$i+1)-$row) & 0xffff)== $this->__get2($share,$i+1)) &&
					((($this->__get1($org,$i+3)-$col) & 0xff)== $this->__get1($share,$i+3))){
					$tmp[$i]=1;
				}
			}
			$i++;
		}
		return $tmp;
	}

	/**
	* change relative token to absolute
	* @access private
	*/
	function __editformula($cur, $formula, $row, $col){
		$i=0;
		$tmp='';
		$lenform=strlen($formula);
		while($i < $lenform){
			if ($cur[$i]){
				$tmp.=chr(0x44);
				$tmp.=pack("v",$this->__get2($formula,$i+1)+$row);
				$tmp.=pack("C",$this->__get1($formula,$i+3)+$col);
				$i+=3;
			} else {
				$tmp.=substr($formula,$i,1);
			}
			$i++;
		}
		return $tmp;
	}

	/**
	* remake Row records
	* @access private
	*/
	function __makeRowRecord($sn){
		$tmp='';
		if(isset($this->rowblock[$sn]))
		foreach((array)$this->rowblock[$sn] as $key => $val) {
			$tmp.=pack("H*",$val['rowhead']);
			$tmp.=pack("vvvv",$key,$val['col1st'],$val['collast'],$val['height']);
			$tmp.=pack("vvvv",$val['notused0'],$val['notused1'],$val['opt0'],$val['opt1']);
		}
		return $tmp;
	}

	/**
	* remake Column records
	* @access private
	*/
	function __makeColRecord($sn){
		$tmp='';
		if(isset($this->colblock[$sn]))
		foreach((array)$this->colblock[$sn] as $key => $val) {
			if ($val['all']){
				$tmp.=$val['all'];
			} else {
				$tmp.=$val['head'];
				$tmp.=pack("vvv",$val['colst'],$val['colen'],$val['width']);
				$tmp.=pack("vvv",$val['xf'],$val['opt'],$val['unk']);
			}
		}
//print bin2hex($tmp)."\n";exit;
		return $tmp;
	}

	/**
	* remake Cell records
	* @access private
	*/
	function __makeCellRecord($sn){
		$tmp='';
		if(isset($this->cellblock[$sn]))
		foreach((array)$this->cellblock[$sn] as $keyR => $rowval) {
			ksort($rowval);
			foreach($rowval as $keyC => $cellval) {
			  if (isset($this->rmcells[$sn][$keyR][$keyC])) continue;
// FIXME		  if ($this->rmcells[$sn][$keyR][$keyC]) continue;
		if (!isset($cellval['record'])) $cellval['record']='';
			  if ($cellval['record']) {
				$tmp.=pack("H*",$cellval['record']);
			  } else {
				$tmp.=pack("vv",$cellval['type'],strlen(pack("H*",$cellval['dat']))+6);
				$tmp.=pack("vvv",$keyR,$keyC,$cellval['xf']);
				$tmp.=pack("H*",$cellval['dat']);
			  }
				if (isset($cellval['sharedform'])) $tmp.=$cellval['sharedform'];
				if (isset($cellval['string'])) $tmp.=$cellval['string'];
			}
		}
		return $tmp;
	}

	/**
	* remake sheet-block
	* @access private
	*/
	function __makeSheet($sn,$ref){
		$this->makeMergeinfo($sn);
		$sno=$this->stable[$sn];
		$tmp='';
		$tmp.=$this->sheetbin[$sno]['preHF'];
		if (isset($this->revise_dat['header'][$sn])){
			$tmp.=pack("vv",Type_HEADER,strlen($this->revise_dat['header'][$sn]));
			$tmp.=$this->revise_dat['header'][$sn];
		} else
		$tmp.=$this->sheetbin[$sno]['header'];
		if (isset($this->revise_dat['footer'][$sn])){
			$tmp.=pack("vv",Type_FOOTER,strlen($this->revise_dat['footer'][$sn]));
			$tmp.=$this->revise_dat['footer'][$sn];
		} else
		$tmp.=$this->sheetbin[$sno]['footer'];
// 2007.04.15 change start by ume
		$tmp.=$this->sheetbin[$sno]['preBT'];
		$tmp.=$this->__makeColRecord($sn);
// 2007.04.15 change end
		$tmp.=$this->sheetbin[$sno]['preCB'];
		$tmp.=$this->__makeRowRecord($sn);
		$tmp.=$this->__makeCellRecord($sn);
// TEST
$tmp.=$this->_makeImageOBJ($sn);
		if ($sn == $sno) {
			$tmp.=$this->sheetbin[$sno]['preMG'];
			$tmp.=$this->makemergecells($sn);
			$tmp.=$this->sheetbin[$sno]['tail'];
		} else {
			if ($this->opt_ref3d){
				$search='5110130001020000b0000b003b....';
				$change='5110130001020000b0000b003b'.bin2hex(pack("v",$ref));
				$this->sheetbin[$sno]['preMG']=pack("H*",ereg_replace($search,$change,bin2hex($this->sheetbin[$sno]['preMG'])));
				$this->sheetbin[$sno]['tail']=pack("H*",ereg_replace($search,$change,bin2hex($this->sheetbin[$sno]['tail'])));
			}
			$tmp.=$this->resetSelectFlag($this->sheetbin[$sno]['preMG']);
			$tmp.=$this->makemergecells($sn);
			$tmp.=$this->resetSelectFlag($this->sheetbin[$sno]['tail']);
		}
		if (isset($this->hlink[$sn]))
		foreach((array)$this->hlink[$sn] as $val){
			$tmp.=$val;
		}
		$tmp.=pack("H*","0a000000");
		return $tmp;
	}

	/**
	* rebuild MERGEDCELLS-record
	* @access private
	*/
	function makemergecells($sn){
		if (! isset($this->mergecells[$sn])) return '';
		if (count($this->mergecells[$sn])==0) return '';
		$ret='';
		$i=0;
		$tmp='';
		foreach($this->mergecells[$sn] as $val){
			$tmp.=pack("v",$val['rows']);
			$tmp.=pack("v",$val['rowe']);
			$tmp.=pack("v",$val['cols']);
			$tmp.=pack("v",$val['cole']);
			if (++$i >=1026 ){
				$ret.=pack("vv",Type_MERGEDCELLS,strlen($tmp)+2).pack("v",1026).$tmp;
				$tmp='';
				$i=0;
			}
		}
		if ($i>0) $ret.=pack("vv",Type_MERGEDCELLS,strlen($tmp)+2).pack("v",$i).$tmp;
		return $ret;
	}

	/**
	* Clear Selected Flag from WINDOW2-record
	* @access private
	*/
	function resetSelectFlag($dat){
		$spos=0;
		$limit=strlen($dat);
		while($spos < $limit){
			$code=$this->__get2($dat,$spos);
			if ($code == Type_WINDOW2){
				$chdat  = substr($dat, 0, $spos+5);
				$chdat .= pack("C", $this->__get1($dat, $spos + 5) & 0xf9);
				$chdat .= substr($dat, $spos + 6);
				return $chdat;
			}
			$spos += $this->__get2($dat,$spos + 2) + 4;
		}
		return $dat;
	}

	/**
	* parse sst-record
	* @access private
	*/
	function __parsesst(&$dat, $pos, $length) {
		$numref=$this->__get4($dat,$pos+8);
		$sspos =12;
		$sstnum=0;
		$limit=$pos + $length +4;
        
		while ($sstnum < $numref) {
			if ($pos+$sspos+2 > $limit) {
				if ($this->__get2($dat,$limit) == Type_CONTINUE) {
					$pos = $limit;
					$length = $this->__get2($dat,$pos + 2);
					$limit += $length + 4;
					$sspos = 4;
				} else break;
			}
			$slen=$this->__get2($dat,$pos+$sspos);
			$tempsst['len']=$slen;
			$opt=$this->__get1($dat,$pos+$sspos+2);
			$sspos += 3;
			if ($opt & 0x01) $slen *=2;
			if ($opt & 0x04) $optlen =4; else $optlen =0;
			if ($opt & 0x08) {
				$optlen +=2;
				$rtnum = $this->__get2($dat,$pos+$sspos);
				if ($opt & 0x04) $apnum = $this->__get4($dat,$pos+$sspos+2);
				else $apnum = 0;
			} else {
				$rtnum = 0;
				if ($opt & 0x04) $apnum = $this->__get4($dat,$pos+$sspos);
				else $apnum = 0;
			}
			$tempsst['opt']=$opt;
			$tempsst['rtn']=$rtnum;
			$tempsst['apn']=$apnum;
			$sspos += $optlen;
			if ($pos+$sspos+$slen > $limit) {
				$fusoku=($pos+$sspos+$slen)-$limit;
				$slen -= $fusoku;
				$sststr=$this->__to_utf16(substr($dat,$pos+$sspos,$slen),$opt);
				if ($opt & 0x01) $fusoku /=2;
				while ($fusoku >0 ) {
					if ($this->__get2($dat,$pos + $length + 4) == Type_CONTINUE) {
						$pos += $length +4;
						$length = $this->__get2($dat,$pos + 2);
						$opt = $this->__get1($dat,$pos + 4);
						$limit = $pos + $length + 4;
						$sspos = 5;
						if ($opt == 1) $fusoku *= 2;
						if ($pos + $sspos + $fusoku > $limit) {
							$fusoku = ($pos + $sspos+ $fusoku) - $limit;
							$sststr.=$this->__to_utf16(substr($dat,$pos + $sspos,$limit-($pos + $sspos)),$opt);
							if ($opt & 0x01) $fusoku /=2;
						} else {
							$sststr.=$this->__to_utf16(substr($dat,$pos + $sspos,$fusoku),$opt);
							$sspos += $fusoku;
							$fusoku=0;
						}
					} else break 2;
				}
			} else {
				$sststr=$this->__to_utf16(substr($dat,$pos+$sspos,$slen),$opt);
				$sspos += $slen;
			}
			if ($rtnum) {
				if ($pos+$sspos+4*$rtnum > $limit) {
					$fusoku=($pos+$sspos+4*$rtnum)-$limit;
					$rt=substr($dat,$pos+$sspos,4*$rtnum - $fusoku);
					if ($this->__get2($dat,$pos + $length + 4) == Type_CONTINUE) {
						$pos += $length + 4;
						$length =$this->__get2($dat,$pos + 2);
						$limit = $pos + $length + 4;
						$sspos = 4;
						$rt.=substr($dat,$limit + $sspos, $fusoku);
						$sspos += $fusoku;
					} else break;
				} else {
					$rt=substr($dat,$pos+$sspos,4*$rtnum);
					$sspos +=4*$rtnum;
				}
			} else $rt="";
			if ($apnum) {
				if ($pos+$sspos+$apnum > $limit) {
					$fusoku=$pos+$sspos+$apnum-$limit;
					$ap=substr($dat,$pos+$sspos,$apnum-$fusoku);
					if ($this->__get2($dat,$limit) == Type_CONTINUE) {
//						$pos = $limit;
						$pos += $length + 4;
						$length = $this->__get2($dat,$pos + 2);
//						$limit += $length + 4;
						$limit = $pos + $length + 4;
						$sspos = 4;
						$ap.=substr($dat,$pos + $sspos, $fusoku);
						$sspos += $fusoku;
					} else break;
				} else {
					$ap=substr($dat,$pos+$sspos,$apnum);
					$sspos +=$apnum;}
			} else $ap="";
//			$sspos +=$apnum;
			$tempsst['str']=bin2hex($sststr);
			$tempsst['rt']=bin2hex($rt);
			$tempsst['ap']=bin2hex($ap);
			$sstarray[$sstnum]=$tempsst;
			$sstnum++;
		}
//print_r($sstarray);
//exit;
		return @$sstarray;
	}

	/**
	* convert charset to UTF16
	* @param $str:string,$opt:0=ascii,1=UTF-16
	* @return UTF16 string
	* @access private
	*/
	function __to_utf16($str,$opt=0)
	{
		return ($opt & 0x01) ? $str : mb_convert_encoding($str, "UTF-16LE", "ASCII");
	}

	/**
	* convert 1,2,4 bytes string to number
	* @param $d:string,$p:position
	* @return number
	* @access private
	*/
	function __get4(&$d, $p) {
		$x = ord($d[$p]) | (ord($d[$p+1]) << 8) |
			(ord($d[$p+2]) << 16) | (ord($d[$p+3]) << 24);
		if ($x > 0x7FFFFFFF) return $x - 0x100000000;
		else return $x;
	}

	/**
	* @access private
	*/
	function __get2(&$d, $p) {
		return ord($d[$p]) | (ord($d[$p+1]) << 8);
	}

	/**
	* @access private
	*/
	function __get1(&$d, $p) {
		return ord($d[$p]);
	}

	/**
	* remake sst record
	* @access private
	*/
	function __makesst(&$sstarray,$totalref) {
		$numref = count($sstarray);
		if (!$numref) return;
		$sstbin='';
		$record = 0xfc;
		$rdat = pack("VV",$totalref,$numref);
		$nokori = 0x2020 - 8;
		foreach ($sstarray as $val) {
			$str=pack("H*",$val['str']);
			$strutf8 = mb_convert_encoding($str, 'utf-8',"UTF-16LE");
			if ($val['rt']) $rt=pack("H*",$val['rt']); else $rt='';
			if ($val['ap']) $ap=pack("H*",$val['ap']); else $ap='';
	
			if ($nokori < 10) {
				$sstbin .= pack("vv",$record,strlen($rdat)).$rdat;
				$record = 0x3c;
				$rdat = '';
				$nokori = 0x2020;
			}
			if (mb_detect_encoding($strutf8,"ASCII,ISO-8859-1")=="ASCII"){
				$opt =0;
				$str = $strutf8;
				$len = strlen($str);
				$lenb = $len;
			} else {
				$opt =1;
				$len = mb_strlen ($str,"UTF-16LE");
				$lenb = 2 * $len;
			}
			if ($ap){
				$opt |= 0x04;
				$apn = strlen($ap);
			} else $apn = 0;
			if ($rt){
				$opt |= 0x08;
				$rtn = strlen($rt) / 4;
			} else $rtn=0;
			$rdat.=pack("vC",$len,$opt);
			if ($rtn) $rdat.=pack("v",$rtn);
			if ($apn) $rdat.=pack("V",$apn);
			$nokori = 0x2020 - strlen($rdat);
			while ($nokori < $lenb) {
				$nokori &= 0xfffe;
				$rdat .= substr($str,0,$nokori);
				$str = substr($str,$nokori);
				$sstbin .= pack("vv",$record,strlen($rdat)).$rdat;
				$lenb -= $nokori;
				$record =0x3c;
				$opt &=1;
				$rdat = pack("C",$opt);
				$nokori = 0x201f;
			}
			$rdat .= $str;
			$nokori = 0x2020 - strlen($rdat);
			while ($nokori < $rtn) {
				$rdat .= substr($rt,0,$nokori);
				$rt = substr($rt,$nokori);
				$sstbin .= pack("vv",$record,strlen($rdat)).$rdat;
				$rtn -= $nokori;
				$record =0x3c;
				$nokori = 0x2020;
				$rdat = '';
			}
			$rdat .= $rt;
			$nokori = 0x2020 - strlen($rdat);
			while ($nokori < $apn) {
				$rdat .= substr($ap,0,$nokori);
				$ap = substr($ap,$nokori);
				$sstbin .= pack("vv",$record,strlen($rdat)).$rdat;
				$apn -= $nokori;
				$record =0x3c;
				$nokori = 0x2020;
				$rdat = '';
			}
			$rdat .= $ap;
			$nokori = 0x2020 - strlen($rdat);
		}
		if ($rdat) $sstbin .= pack("vv",$record,strlen($rdat)).$rdat;
		return $sstbin;
	}

	/**
	* Parse Excel file
	* @param  $filename:full path for OLE file
	* @access public
	* @example ./sample_ex1.php sample_ex1
	*/
	function parseFile($filename,$mode=null){
		if ($mode == 1) $this->opt_parsemode = 1;
	if ($filename !=''){
		$dat = $this->__oleread($filename);
		if ($this->isError($dat)) return $dat;
		if (strlen($dat) < 256) {
			return $this->raiseError("Contents is too small (".strlen($dat).")\nProbably template file is not Excel file.\n");
		}
	} else {
		$dat=pack("H*","09081000000605002d20cd07c9c0000006030000e1000200b004c10002000000e20000005c007000210000457863656c5f526576697365722020687474703a2f2f6368617a756b652e636f6d2020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202042000200b0043d001200e0011e001b2175123800000000000100580222000200000031002000dc000000ff7f900100000003806a08012dff33ff200030ffb430b730c330af3031002000dc000000ff7f900100000003806a08012dff33ff200030ffb430b730c330af3031002000dc000000ff7f900100000003806a08012dff33ff200030ffb430b730c330af3031002000dc000000ff7f900100000003806a08012dff33ff200030ffb430b730c330af303100200078000000ff7f900100000003806a08012dff33ff200030ffb430b730c330af30e000140000000000f5ff200000000000000000000000c020e000140001000000f5ff200000f40000000000000000c020e000140001000000f5ff200000f40000000000000000c020e000140002000000f5ff200000f40000000000000000c020e000140002000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e000140000000000f5ff200000f40000000000000000c020e0001400000000000100200000000000000000000000c02093020400108005ff93020400118006ff93020400128003ff93020400138007ff93020400148004ff93020400008000ff85000e004603000000000600536865657431fc0008000000000000000000ff00020008000a00000009081000000610002d20cd07c9c00000060300000b021400000000000200000004000000b80d00000c0e00000d00020001000c00020064000f00020001002502040000000e0114000b00080000264368656164657215000b000800002643666f6f746572830002000000840002000000a10022000900640001000100010002002c012c01fca9f1d24d62e03ffca9f1d24d62e03f010055000200080000020e000200000004000000000001000000080210000200000001000e010000000000010f00010206000200000015003e021200b606000000004000000000000000000000000a000000");
	}
		$presheet=1;
		$pos = 0;
		$version = $this->__get2($dat,$pos + 4);
		$substreamType = $this->__get2($dat,$pos + 6);
		if ($version != Code_BIFF8) {
			return $this->raiseError("Contents is not BIFF8 format.\n");
		}
		if ($substreamType != Code_WorkbookGlobals) {
			return $this->raiseError("Contents is unknown format.\nCan't find WorkbookGlobal.");
		}
		$code=-1;
		$poslimit=strlen($dat);
		while ($code != Type_EOF){
			if ($pos > $poslimit){
				return $this->raiseError("Global Area Read Error\nTemplate file is broken");
			}
		    $code = $this->__get2($dat,$pos);
		    $length = $this->__get2($dat,$pos+2);
		    switch ($code) {
			case Type_BOF:
				$this->wbdat = substr($dat, $pos, $length+4);
				if ($this->PW_Hash_write!==NULL && $this->PW_Hash_write > 0){
					$this->wbdat .= pack("H*","86000000");
				}
			    break;
			case Type_FILEPASS:
				if (strlen($this->openFPass)<1) return $this->raiseError("Set File-Open-Password. \nThis file is protected.");
				if (substr($dat,$pos+4,6)!=pack("H*","010001000100")) return $this->raiseError("It doesn't correspond to the encryption of this type.");
				$docid=substr($dat,$pos+10,16);
				$salt=substr($dat,$pos+26,16);
				$hashedsalt=substr($dat,$pos+42,16);
				$size=strval(strlen($dat));
				$md5= NEW XL_Crypto();
				if (!$md5->verifypwd($this->openFPass, $docid, $salt, $hashedsalt))
					$this->raiseError("Wrong Password is setted.\n");;
				$op = $md5->makeXorPtn($size);
				if (strlen($op) < 128) return $this->raiseError($op);
				$datx=$dat ^ $op;
				if (strlen($dat) <= strlen($datx)){
					$dat=$this->decrypt($dat,$datx);
				} else {
					return $this->raiseError("abnormal hash. \nThis file is protected.");
				}
				break;
			case Type_SST:
				$this->globaldat['presst']=$this->wbdat;
				$this->wbdat='';
				$this->eachsst = $this->__parsesst($dat, $pos, $length);
				while ($this->__get2($dat,$pos + $length + 4) == Type_CONTINUE){
					$pos += $length + 4;
					$length = $this->__get2($dat,$pos+2);
				}
			    break;
			case Type_EXTSST:
//				$this->globaldat['exsstbin'] = '';	// FIXME
				break;
			case Type_OBJPROJ:
			case Type_BUTTONPROPERTYSET:
				break;
			case Type_BOUNDSHEET:
				if ($presheet) {
					$this->globaldat['presheet']=$this->wbdat;
					$this->wbdat='';
					$presheet=0;
				}
				$rec_offset = $this->__get4($dat, $pos+4);
			    $sheetno['code'] = substr($dat, $pos, 2);
			    $sheetno['length'] = substr($dat, $pos+2, 2);
			    $sheetno['offsetbin'] = substr($dat, $pos+4, 4);
			    $sheetno['offset'] = $rec_offset;
			    $sheetno['visible'] = substr($dat, $pos+8, 1);
			    $sheetno['type'] = substr($dat, $pos+9, 1);
			    $sheetno['name'] = substr($dat, $pos+10, $length-6);
			    $this->boundsheets[] = $sheetno;
			    break;
			case Type_COUNTRY:
					$this->wbdat .= substr($dat, $pos, $length+4);
					$this->globaldat['presup']=$this->wbdat;
					$this->wbdat='';
			    break;
			case Type_XCT:
			case Type_CRN:
				break;
			case Type_SUPBOOK:
// tentative countermeasures for unknown SUPBOOK-record(External References) on 2007.1.29
//
				if (substr($dat, $pos+6,2)=="\x01\x04"){
					$this->globaldat['presup'].=$this->wbdat;
					$this->globaldat['supbook']=substr($dat, $pos, $length+4);
				}
				$this->wbdat='';
				break;
			case Type_EXTERNSHEET:
				if (strlen($this->globaldat['presup'])==0) {
					$this->globaldat['presup'].=$this->wbdat;
					$this->wbdat='';
				}
				$this->globaldat['extsheet']=substr($dat, $pos, $length+4);
				$this->globaldat['name']='';
				while($this->__get2($dat, $pos+$length+4)==Type_NAME){
					$pos +=$length+4;
					$length = $this->__get2($dat,$pos+2);
					if ($this->__get2($dat,$pos+4)!=0x20 || $this->__get1($dat,$pos+11)!=0 || $this->__get1($dat,$pos+12)==0){
//						$this->globaldat['name'].=substr($dat, $pos, $length+4);
					} else {
						$this->globaldat['namerecord'].=substr($dat, $pos, $length+4);
						$lenform=$this->__get2($dat,$pos+8);
						$namtype=$this->__get1($dat,$pos+19);
						$tmp['flags2notu']=substr($dat, $pos, 12);
						$tmp['sheetindex']=$this->__get2($dat,$pos+12);
						$tmp['menu2name']=substr($dat, $pos+14, 6);
						$tmp['formula']=$this->analizeform(substr($dat,$pos+20,$lenform));
						$tmp['remain']=substr($dat,$pos+20+$lenform,$length-(16+$lenform));
						$this->boundsheets[$this->__get2($dat,$pos+12)-1]['namerecord'][$namtype]=$tmp;
					}
				}
			    break;
			case Type_WRITEACCESS:
/*
				$wa = "5c007000120000687474703a2f2f6368617a756b652e636f6d";
				$wa.= "20202020202020202020202020202020202020202020202020";
				$wa.= "20202020202020202020202020202020202020202020202020";
				$wa.= "20202020202020202020202020202020202020202020202020";
				$wa.= "20202020202020202020202020202020";
				$this->wbdat .= pack("H*",$wa);
*/
				$ERVer= mb_convert_encoding(Reviser_Version,'UTF-16LE','auto');
				$wa = "\x5c\x00\x70\x00" . pack("C", 34 + mb_strlen($ERVer,'UTF-16LE'));
				$wa.= pack("H*","000145007800630065006c005f0052006500760069007300650072002000");
				$wa.= $ERVer;
				$wa.= pack("H*","2000200068007400740070003a002f002f006300680061007a0075006b0065002e0063006f006d00");
				$wa = str_pad($wa, 0x74);
				$this->wbdat .=$wa;
				if ($this->PW_Hash_write!==NULL && $this->PW_Hash_write > 0){
					$this->wbdat.= "\x5b\x00\x28\x00";
					$this->wbdat.=pack("vv",$this->Flag_Read_Only,$this->PW_Hash_write);
					$this->wbdat.= pack("H*","210000457863656c5f526576697365722020687474703a2f2f6368617a756b652e636f6d");
				} else if ($this->Flag_Read_Only == 1){
					$this->wbdat.= pack("H*","5b000600010000000000");
				}
			    break;
			case Type_WRITEPROT:
				if ($this->PW_Hash_write===NULL){
					$this->wbdat .= substr($dat, $pos, $length+4);
				}
			    break;
			case Type_FILESHARING:
				if ($this->__get2($dat,$pos+6)!==$this->PW_Hash_read){
					return $this->raiseError("Template-file is read-only. Set correct-password.");
					break;
				}
				if ($this->PW_Hash_write===NULL){
					$this->wbdat .= substr($dat, $pos, $length+4);
				}
			    break;
			case Type_EOF:
				$this->globaldat['last']= $this->wbdat . substr($dat, $pos, $length+4);
			    break;

			case Type_FONT:
			case Type_FORMAT:
			case Type_XF:
				$this->wbdat .= substr($dat, $pos, $length+4);
				if ($this->opt_parsemode) $this->saveAttrib($code,substr($dat, $pos, $length+4));
				break;
			default:
				$this->wbdat .= substr($dat, $pos, $length+4);
			}
			$pos += $length + 4;
		}
		foreach ($this->boundsheets as $key=>$val){
		    $res = $this->__parsesheet($dat,$key,$val['offset']);
			if ($this->isError($res)) return $res;
		}
	}

	/**
	* Remake Excel file
	* @param  $filename:file-name for web output
	* @return stdout for web-output
	* @access private
	*/
	function makeFile($filename,$path=null){
		$this->_makesupblock();
		$totalref = count($this->eachsst);	// FIXME
		$sstbin=$this->__makesst($this->eachsst,$totalref);
		$tmplen=strlen($this->globaldat['presheet']);
		$tmplen += strlen($this->globaldat['presst']);
		$tmplen += strlen($this->globaldat['last']);
		$tmplen += strlen($this->globaldat['presup']);
		$tmplen += strlen($this->globaldat['supbook']);
		$tmplen += strlen($this->globaldat['extsheet']);
		$tmplen += strlen($this->globaldat['name']);
		$tmplen += strlen($this->globaldat['namerecord']);
		$tmplen += strlen($sstbin.$this->globaldat['exsstbin']);
//		$refnum1=$refnum;
		foreach ($this->boundsheets as $key=>$val){
			$tmplen += strlen($val['code']);
			$tmplen += strlen($val['length']);
			$tmplen += strlen($val['offsetbin']);
			$tmplen += strlen($val['visible']);
			$tmplen += strlen($val['type']);
			$tmplen += strlen($val['name']);
//			$sheetdat[$key]=$this->__makeSheet($key,$refnum1++);
			$sheetdat[$key]=$this->__makeSheet($key,$key);
		}
	
		foreach ((array)$sheetdat as $key=>$val){
			$this->boundsheets[$key]['offsetbin']=pack("V",$tmplen);
			$tmplen += strlen($val);
		}
	// make global-block
		$tmp=$this->globaldat['presheet'];
		foreach ($this->boundsheets as $key=>$val){
			$tmp .= $val['code'];
			$tmp .= $val['length'];
			$tmp .= $val['offsetbin'];
			$tmp .= $val['visible'];
			$tmp .= $val['type'];
			$tmp .= $val['name'];
		}
		$tmp .= $this->globaldat['presup'].$this->globaldat['supbook'];
		$tmp .= $this->globaldat['extsheet'];
		$tmp .= $this->globaldat['name'];
		$tmp .= $this->globaldat['namerecord'];
		$tmp .= $this->globaldat['presst'] . $sstbin . $this->globaldat['exsstbin'];
		$tmp .= $this->globaldat['last'];
		foreach ((array)$sheetdat as $val){
			$tmp .= $val;
		}

	// If password is set
		if (strlen($this->saveFPass)>0){
/*
echo strlen($tmp);
echo "\n";
echo bin2hex(substr($tmp,0,80));
exit;
*/
			$res=$this->encrypt($tmp);
			if ($this->isError($res)) return $res;
		}
	// from here making Excel-file
		if (($path === null) || (trim($path)=="")) {
			header("Content-type: application/vnd.ms-excel");
			header("Content-Disposition: attachment; filename=\"$filename\"");
			header("Expires: 0");
			header("Cache-Control: must-revalidate, post-check=0,pre-check=0");
			header("Pragma: public");
			print $this->makeole2($tmp);
		} else {
			if (substr($path,-1) == '/') $path = substr($path,0,-1);
			if (!file_exists($path)) return $this->raiseError("The path $path does not exist.");
			$filename = $path . '/' . $filename;
			$_FILEH_ = @fopen($filename, "wb");
			if ($_FILEH_ == false) {
				return $this->raiseError("Can't open $filename. It may be in use or protected.");
			}
			fwrite($_FILEH_, $this->makeole2($tmp));
			@fclose($_FILEH_);
		}
	}

	/**
	* Remake Cell block
	* @access private
	*/
	function reviseCell(){
		$shname=array();
		$tmpsn=count($this->boundsheets);
		for ($i=0;$i<$tmpsn;$i++){
			$shname[$this->boundsheets[$i]['name']] = 0;
		}
		for ($i=0;$i<$tmpsn;$i++){
			$this->stable[$i]=$i;
			$shname[$this->boundsheets[$i]['name']]++;
		}

		foreach($this->dupsheet as $val){
			if (isset($this->boundsheets[$val['orgsn']])){
				for($i=0;$i<$val['count'];$i++){
					$this->stable[$tmpsn+$i]=$val['orgsn'];
				  if (isset($this->mergecells[$val['orgsn']]))
					$this->mergecells[$tmpsn+$i]=$this->mergecells[$val['orgsn']];
				  if (isset($this->rowblock[$val['orgsn']]))
					$this->rowblock[$tmpsn+$i]=$this->rowblock[$val['orgsn']];
				  if (isset($this->colblock[$val['orgsn']]))
					$this->colblock[$tmpsn+$i]=$this->colblock[$val['orgsn']];
				  if (isset($this->cellblock[$val['orgsn']]))
					$this->cellblock[$tmpsn+$i]=$this->cellblock[$val['orgsn']];
					$this->boundsheets[$tmpsn+$i]=$this->boundsheets[$val['orgsn']];
					if (isset($this->revise_dat['sheetname'][$tmpsn+$i])){
						$this->boundsheets[$tmpsn+$i]['name'] = $this->revise_dat['sheetname'][$tmpsn+$i];
						$this->boundsheets[$tmpsn+$i]['length'] = pack("v",6 + strlen($this->revise_dat['sheetname'][$tmpsn+$i]));
						if (isset($shname[$this->boundsheets[$tmpsn+$i]['name']])){
							$shname[$this->boundsheets[$tmpsn+$i]['name']] +=1;
						} else  $shname[$this->boundsheets[$tmpsn+$i]['name']] =1;
					} else {
	//	if the names are same, add different-number
						if ($shname[$this->boundsheets[$tmpsn+$i]['name']] > 0){
							$shname[$this->boundsheets[$tmpsn+$i]['name']]++;
							$dupstr = '('.($shname[$this->boundsheets[$tmpsn+$i]['name']] -1).')';
							$strcnt=$this->__get1($this->boundsheets[$tmpsn+$i]['name'],0) + strlen($dupstr);
							if ($this->__get1($this->boundsheets[$tmpsn+$i]['name'],1) == 0) {
								$this->boundsheets[$tmpsn+$i]['name'] .= $dupstr;
							} else {
								$this->boundsheets[$tmpsn+$i]['name'] .= mb_convert_encoding($dupstr, "UTF-16LE", "ASCII");
							}
							$this->boundsheets[$tmpsn+$i]['name']=pack("C",$strcnt).substr($this->boundsheets[$tmpsn+$i]['name'],1);
							$this->boundsheets[$tmpsn+$i]['length'] = pack("v",6 + strlen($this->boundsheets[$tmpsn+$i]['name']));
						}
					}
				}
				$tmpsn += $val['count'];
			}
		}
		foreach($this->boundsheets as $key=>$val){
			if (isset($this->revise_dat['sheetname'][$key]))
			if (strlen($this->revise_dat['sheetname'][$key])){
			    $this->boundsheets[$key]['name'] = $this->revise_dat['sheetname'][$key];
			    $this->boundsheets[$key]['length'] = pack("v",6 + strlen($this->revise_dat['sheetname'][$key]));
			}
		}
// end of sheet dup  TEST 2009.02.15
		if (isset($this->revise_dat['copy_row'])){
			$sheetnum=count($this->boundsheets);
			foreach((array)$this->revise_dat['copy_row'] as $key => $val) {
				if ($val['sheetsrc']>=$sheetnum) continue;
				if ($val['sheetdest']>=$sheetnum) continue;
				if (($val['num']<1) || ($val['num']>65355)) continue;
				for($i=0;$i<$val['num'];$i++){
					$j=($val['inc']==0)? 0:$i;
					if (isset($this->cellblock[$val['sheetsrc']][$val['rowsrc']+$j])){
					$this->cellblock[$val['sheetdest']][$val['rowdest']+$i]=$this->cellblock[$val['sheetsrc']][$val['rowsrc']+$j];
					foreach($this->cellblock[$val['sheetdest']][$val['rowdest']+$i] as $ckey=>$cval){
						$this->cellblock[$val['sheetdest']][$val['rowdest']+$i][$ckey]['record']='';
						}
					}
					if (isset($this->rowheight[$val['sheetsrc']][$val['rowsrc']+$j]))
					$this->rowheight[$val['sheetdest']][$val['rowdest']+$i]=$this->rowheight[$val['sheetsrc']][$val['rowsrc']+$j];
				}
			}
		}

		if (isset($this->revise_dat['copy_col'])){
			$sheetnum=count($this->boundsheets);
			foreach((array)$this->revise_dat['copy_col'] as $key => $cval) {
				if ($cval['sheetsrc']>=$sheetnum) continue;
				if ($cval['sheetdest']>=$sheetnum) continue;
				if (($cval['num']<1) || ($cval['num']>255)) continue;
				foreach($this->cellblock[$cval['sheetsrc']] as $rkey=>$rval){
					for($i=0;$i<$cval['num'];$i++){
						$j=($cval['inc']==0)? 0:$i;
						if (isset($this->cellblock[$cval['sheetsrc']][$rkey][$cval['colsrc']+$j])){
						$this->cellblock[$cval['sheetdest']][$rkey][$cval['coldest']+$i]=$this->cellblock[$cval['sheetsrc']][$rkey][$cval['colsrc']+$j];
						$this->cellblock[$cval['sheetdest']][$rkey][$cval['coldest']+$i]['record']='';
						}
					}
				}
				for($i=0;$i<$cval['num'];$i++){
					$j=($cval['inc']==0)? 0:$i;
					if (isset($this->colwidth[$cval['sheetsrc']][$cval['colsrc']+$j]))
					$this->colwidth[$cval['sheetdest']][$cval['coldest']+$i]=$this->colwidth[$cval['sheetsrc']][$cval['colsrc']+$j];
				}
			}
		}
// TEST end

		if(isset($this->revise_dat['replace']))
		foreach((array)$this->revise_dat['replace'] as $val) {
			$search=bin2hex($val['org']);
			$replace=bin2hex($val['new']);
			foreach((array)$this->eachsst as $key => $dmy) {
				$this->eachsst[$key]['str']=str_replace($search, $replace, $this->eachsst[$key]['str']);
			}
		}

// Start of Experimental code on 2007/10/13
		if (isset($this->revise_dat['option']))
		foreach((array)$this->revise_dat['option'] as $key => $val) {
			if (($this->__get2($val['record'],2)+4) != strlen($val['record'])) continue;
			if ($val['refsheet'] === null) $val['refsheet'] = $val['sheet'];
			if (($val['refrow'] !== null) && ($val['refcol'] !== null)) {
				$xf= (isset($this->cellblock[$val['refsheet']][$val['refrow']][$val['refcol']]['xf'])) ? $this->cellblock[$val['refsheet']][$val['refrow']][$val['refcol']]['xf'] : $this->_getcolxf($val['refsheet'],$val['refcol']);
			} else {
				$xf= (isset($this->cellblock[$val['sheet']][$val['row']][$val['col']]['xf'])) ? $this->cellblock[$val['sheet']][$val['row']][$val['col']]['xf'] : $this->_getcolxf($val['sheet'],$val['col']);
			}
			$data = substr($val['record'],0,8).pack('v', $xf).substr($val['record'],10);
			$this->cellblock[$val['sheet']][$val['row']][$val['col']]['record']=bin2hex($data);
		}
// End of Experimental code

		if (isset($this->revise_dat['add_str']))
		foreach((array)$this->revise_dat['add_str'] as $key => $val) {
			$this->_addString($val['sheet'],$val['row'], $val['col'], $val['str'], $val['refrow'], $val['refcol'], $val['refsheet']);
		}
		if (isset($this->revise_dat['add_num']))
		foreach((array)$this->revise_dat['add_num'] as $key => $val) {
			$this->_addNumber($val['sheet'],$val['row'], $val['col'], $val['num'], $val['refrow'], $val['refcol'], $val['refsheet']);
		}
		if (isset($this->revise_dat['add_blank']))
		foreach((array)$this->revise_dat['add_blank'] as $key => $val) {
			$this->_addBlank($val['sheet'],$val['row'], $val['col'], $val['refrow'], $val['refcol'], $val['refsheet']);
		}

		if (isset($this->revise_dat['add_formula']))
		foreach((array)$this->revise_dat['add_formula'] as $key => $val) {
			$res=$this->_addFormula($val['sheet'],$val['row'], $val['col'], $val['formula'], $val['refrow'], $val['refcol'], $val['refsheet'],$val['opt']);
			if ($this->isError($res)) return $res;
		}
		if (count((array)$this->colwidth)>0)
		foreach((array)$this->colwidth as $key => $val) {
			foreach($val as $key1 => $val1) {
				if (isset($this->colblock[$key][$key1])){
					if ($val1 == 0) {
						$this->colblock[$key][$key1]['opt'] |=0x01;
					} else {
						$this->colblock[$key][$key1]['width']=$val1;
					}
					$this->colblock[$key][$key1]['all']='';
				} else {
					$work['head']=pack("H*","7d000c00");
					$work['colst']=$key1;
					$work['colen']=$key1;
					$work['xf']=$this->_getcolxf($key,$key1);
					$work['unk']=0x02;
					$work['all']='';
					if ($val1 > 0) {
						$work['width']=$val1+ 0x0a0;
						$work['opt']=0x02;
					} else {
						$work['width']=0x0900;
						$work['opt']=0x03;
					}
					$this->colblock[$key][$key1]=$work;
				}
			}
		}

		if (count((array)$this->rowheight)>0)
		foreach((array)$this->rowheight as $key => $val) {
			foreach($val as $key1 => $val1) {
				if (isset($this->rowblock[$key][$key1])){
					if ($val1 == 0) {
						$this->rowblock[$key][$key1]['opt0'] |=0x20;
					} else {
						$this->rowblock[$key][$key1]['height']=$val1;
						$this->rowblock[$key][$key1]['opt0'] |=0x40;
					}
				} else {
					$this->rowblock[$key][$key1]['rowhead']="08021000";
					$this->rowblock[$key][$key1]['col1st']=0;
					$this->rowblock[$key][$key1]['collast']=count($this->cellblock[$key][$key1]);
					$this->rowblock[$key][$key1]['height']=$val1;
					$this->rowblock[$key][$key1]['notused0']=0;
					$this->rowblock[$key][$key1]['notused1']=0;
					$this->rowblock[$key][$key1]['opt0']=0x40;
					$this->rowblock[$key][$key1]['opt1']=0x0f;
				}
			}
		}
		krsort($this->rmsheets);
		$refnum=0;
		foreach ($this->rmsheets as $key => $val) {
			if ((count($this->boundsheets) > 1) && $val){
				unset($this->boundsheets[$key]);
			}
		}
		$this->_setPrintInfo();
	}

	/**
	* make OLE container & output to STDOUT
	* @param $tmpbin:binary data
	* @return web header and data
	* @access private
	*/
    function makeole2(& $tmpbin){
		$naiyou['bin']=$tmpbin;
		$naiyou['name']='Workbook';
		$streams[]=$naiyou;
	if (isset($this->orgStreams["\x05SummaryInformation"])){
		$naiyou['bin']=$this->orgStreams["\x05SummaryInformation"]['dat'];
		$naiyou['name']="\x05SummaryInformation";
		$streams[]=$naiyou;
	}
	if (isset($this->orgStreams["\x05DocumentSummaryInformation"])){
	        $naiyou['bin']=$this->orgStreams["\x05DocumentSummaryInformation"]['dat'];
	        $naiyou['name']="\x05DocumentSummaryInformation";
	        $streams[]=$naiyou;
	}
	$AlTbls=0;
	$blocks=array();
	$MSATSID=array();
	$nextSec=0;
	$rootentry=str_pad($this->asc2utf('Root Entry'), 64, "\x00")	//0- 64
		. pack("v",2*(1+strlen('Root Entry')))		//64- 2
		. "\x05"		//66- 1
		. "\x01"		//67- 1
		. pack("V",-1)	//68- 4
		. pack("V",-1)	//72- 4
		. ((count($streams)==3)? pack("V",2):pack("V",1))	//76- 4
		. str_repeat("\x00", 16)	//80- 16
		. pack("V",0)	//96- 4
		. pack("d",0)	//100- 8
		. pack("d",0)	//108- 8
		. pack("V",0)	//116- 4
		. pack("V",0)	//120- 4
		. pack("V",0);	//124- 4
	foreach($streams as $key=>$dat){
		$orglen=strlen($dat['bin']);
		if ($orglen < 0x1000) {
			$streams[$key]['bin']=str_pad($dat['bin'], 0x1000, "\x00");
			$orglen = 0x1000;
		} else {
			if ($orglen % 512 != 0)
			$streams[$key]['bin'] .= str_repeat("\x00", 512 - ($orglen % 512));
		}
		$needSecs = strlen($streams[$key]['bin'])/512;
		$AlTbls += $needSecs;
	// 1st each binary-dat
		for($i=0;$i<$needSecs-1;$i++){
			$blocks[$nextSec+$i]=$nextSec+$i+1;
		}
		$blocks[$nextSec+$i]=-2;
		$userstream=str_pad($this->asc2utf($dat['name']), 64, "\x00")	//0- 64
			. pack("v",2*(1+strlen($dat['name'])))		//64- 2
			. "\x02"		//66- 1
			. "\x01";		//67- 1
		if (count($streams)==3 && $key==1) {
			$userstream.= pack("VVV", 1, 3,-1);
		} else {
			$userstream.= pack("VVV",-1,-1,-1);	//68-76- 4x3
		}
			$userstream.= str_repeat("\x00", 16)	//80- 16
			. pack("V",0)	//96- 4
			. pack("d",0)	//100- 8
			. pack("d",0)	//108- 8
			. pack("V",$nextSec)	//116- 4
			. pack("V",$orglen)	//120- 4
			. pack("V",0);	//124- 4
		$nextSec=$nextSec+$i+1;
		$rootentry.=$userstream;
	}
	$rootentry.=str_repeat("\x00", 512 - (strlen($rootentry) % 512));
	//2ns RootEntry directory
	$rootSec=$nextSec;
	$DirSecs= strlen($rootentry) / 512;
	for($i=0;$i<$DirSecs-1;$i++){
		$blocks[$nextSec+$i]=$nextSec+$i+1;
	}
	$blocks[$nextSec+$i]=-2;
	$nextSec=$nextSec+$i+1;
	//3rd allocation table
	$alcsecs=floor((count($blocks)+127)/127);
	for($i=0;$i<$alcsecs;$i++){
		$blocks[$nextSec+$i]=-3;
		$MSATSID[]=$nextSec+$i;
	}
	$nextSec=$nextSec+$i+1;
	$blocks[$nextSec+$i]=-2;
	$totalAlTblnum=$alcsecs;
	$head=pack("H*","D0CF11E0A1B11AE1")
			. str_repeat("\x00", 16)
			. pack("v", 0x3b)
			. pack("v", 0x03)
			. pack("v", -2)
			. pack("v", 9)
			. pack("v", 6)
			. str_repeat("\x00", 10)
			. pack("V", $totalAlTblnum)
			. pack("V", $rootSec)
			. pack("V", 0)
			. pack("V", 0x1000)
			. pack("V", 0)  //Short Block Depot
			. pack("V", 1)
			. pack("V", -2)	//$masterAlTbl
			. pack("V", 0);	//$masterAlnum)
	// make OLE container
	$oledat =$head;
	for($i=0;$i<109;$i++){
		if(isset($MSATSID[$i])){
			$oledat.=pack("V",$MSATSID[$i]);
		} else {
			$oledat.=pack("V",-1);
		}
	}
	foreach($streams as $dat){
		$oledat.=$dat['bin'];
	}
	$oledat.=$rootentry;
	for($i=0;$i<$alcsecs*128;$i++){
		if(isset($blocks[$i])){
			$oledat.=pack("V",$blocks[$i]);
		} else {
			$oledat.=pack("V",-1);
		}
	}
	return $oledat;
    }

	/**
	* convert charset ASCII to UTF16
	* @param $ascii string
	* @return UTF16 string
	* @access private
	*/
	function asc2utf($ascii){
		$utfname='';
		for ($i = 0; $i < strlen($ascii); $i++) {
			$utfname.=$ascii{$i}."\x00";
		}
		return $utfname;
	}

	/**
	* get Cell Attribute
	* @access private
	*/
	function saveAttrib($code,$dat){
		switch ($code) {
			case Type_FONT:
				$this->recFONT[]=$dat;
				break;
			case Type_FORMAT:
				$fmt=$this->cnvstring(substr($dat,6),2);
				$this->recFORMAT[$this->__get2($dat,4)]=$fmt;
				break;
			case Type_XF:
				$this->recXF[]=$dat;
				break;
		}
	}

	/**
	* convert string from UTF to internal-charset
	* @access private
	*/
	function cnvstring($chars,$len){
		if ($len==1) {
			$strpos=2;
			$opt=$this->__get1($chars,1);
		} elseif ($len==2){
			$strpos=3;
			$opt=$this->__get1($chars,2);
		} else return substr($chars,2);
		if ($opt)
			return mb_convert_encoding(substr($chars,$strpos),$this->charset,'UTF-16LE');
		else
			return substr($chars,$strpos);
	}

	/**
	* Get Cell Value
	* @param integer $sn sheet number
	* @param integer $row Row position
	* @param integer $col Column position  0indexed
	* @return mixed cell value
	* @access public
	* @example ./sample_ex1.php sample_ex1
	*/
	function getCellVal($sn,$row,$col){
		if (isset($this->cellblock[$sn][$row][$col])) {
			$cell=$this->cellblock[$sn][$row][$col];
			$tmp['type'] = $cell['type'];
			switch ($cell['type']) {
				case Type_LABEL:
					$desc=$this->cnvstring(pack("H*",$cell['dat']),2);
					break;
				case Type_LABELSST:
					$c=pack("H*",$cell['dat']);
					$strnum=$this->__get2($c,0);
					$sstr=$this->eachsst[$strnum]['str'];
					$desc=mb_convert_encoding(pack("H*",$sstr),$this->charset,'UTF-16LE');
					break;
				case Type_RK:
				case Type_RK2:
					$c=pack("H*",$cell['dat']);
					$rknum = $this->__get4($c,0);
					if (($rknum & 0x02) != 0) {
						$value = $rknum >> 2;
					} else {
						$sign = ($rknum & 0x80000000) >> 31;
						$exp = ($rknum & 0x7ff00000) >> 20;
						$mantissa = (0x100000 | ($rknum & 0x000ffffc));
						$value = $mantissa / pow( 2 , (20- ($exp - 1023)));
						if ($sign) {$value = -1 * $value;}
					}
					if (($rknum & 0x01) != 0) $value /= 100;

					$desc=$value;
					break;
				case Type_NUMBER:
					$temp=(pack("N",1)==pack("L",1)) ? strrev(pack("H*",$cell['dat'])) : pack("H*",$cell['dat']);
					$strnum=unpack("d",$temp);
					$desc=$strnum[1];
					break;
				case Type_FORMULA:
				case Type_FORMULA2:
					$result=substr(pack("H*",$cell['dat']),0,8);
					if (substr($result,6,2)=="\xFF\xFF"){
						switch (substr($result,0,1)) {
						case "\x00":
							$desc=$this->cnvstring(substr($cell['string'],4),2);
							break;
						case "\x01":
							$desc=(substr($result,2,1)=="\x01")? "TRUE":"FALSE";
							break;
						case "\x02": $desc='#ERROR!';
							break;
						case "\x03": $desc='';
							break;
						}
					} else {
						$t0=(pack("N",1)==pack("L",1)) ? strrev($result) : $result ;
						$desc0=unpack("d",$t0);
						$desc=$desc0[1];
					}
					break;
				case Type_BOOLERR:
					$result=pack("H*",$cell['dat']);
					if ($this->__get1($result,1) !=0) {
						$desc='#ERROR!';
					} elseif ($this->__get1($result,0) !=0) {
						$desc = "TRUE";
					} else {
						$desc = "FALSE";
					}
					break;
				case Type_BLANK:
					$desc='';
					break;
				default:
					$tmp['type'] = -1;
					$desc='';
			}
			$tmp['val'] = $desc;
		} else {
			$tmp['type'] = 0;
			$tmp['val'] = '';
		}
		return $tmp;
	}

	/**
	* Get Cell Attribute
	* @param integer $sn sheet number
	* @param integer $row Row position
	* @param integer $col Column position
	* @return mixed cell value
	* @access public
	* @example ./sample_ex1.php sample_ex1
	*/
	function getCellAttrib($sn,$row,$col){
		if ($this->opt_parsemode !=1) return -1;
		$xfno=$this->cellblock[$sn][$row][$col]['xf'];
		if ($xfno !== null) {
			$dat=$this->recXF[$xfno];
			$xf['attrib']=($this->__get1($dat,13) & 0xfc) >> 2;
			$xf['stylexf']=($this->__get1($dat,8) & 0x4) >> 2;
			$oya=($this->__get2($dat,8) & 0xfff0) >> 4;
			if ($oya != 0xfff) $xf['parent']=$oya;
			$cond = $xf['stylexf'] ? ~$xf['attrib'] : $xf['attrib'];
			if ($cond & 0x2)
				$xf['fontindex']=$this->__get2($dat,4)-1;
				else $xf['fontindex']=0;
			if ($cond & 0x1)
				$xf['formindex']=$this->__get2($dat,6);
				else $xf['formindex']=0;
//			if ($cond & 0x4){
				$xf['halign']=$this->__get1($dat,10) & 0x7;
				$xf['wrap']=($this->__get1($dat,10) & 0x8) >> 3;
				$xf['valign']=($this->__get1($dat,10) & 0x70)>> 4;
				$xf['rotation']=$this->__get1($dat,11);
//			}
//			if ($cond & 0x8){
				$xf['Lstyle']=$this->__get1($dat,14) & 0x0f;
				$xf['Rstyle']=($this->__get1($dat,14) & 0xf0) >> 4;
				$xf['Tstyle']=$this->__get1($dat,15) & 0x0f;
				$xf['Bstyle']=($this->__get1($dat,15) & 0xf0) >> 4;
				$xf['Lcolor']=$this->__get1($dat,16) & 0x7f;
				$xf['Rcolor']=($this->__get2($dat,16) & 0x3f80) >> 7;
				$xf['diagonalL2R']=($this->__get1($dat,17) & 0x40) >> 6;
				$xf['diagonalR2L']=($this->__get1($dat,17) & 0x80) >> 7;
				$xf['Tcolor']=$this->__get1($dat,18) & 0x7f;
				$xf['Bcolor']=($this->__get2($dat,18) & 0x3f80) >> 7;
				$xf['Dcolor']=($this->__get4($dat,18) & 0x1fc000) >> 14;
				$xf['Dstyle']=($this->__get2($dat,20) & 0x1e0) >> 5;
//			}
//			if ($cond & 0x10){
				$xf['fillpattern']=($this->__get1($dat,21) & 0xfc) >> 2;
				$xf['PtnFRcolor']=$this->__get1($dat,22) & 0x7f;
				$xf['PtnBGcolor']=($this->__get2($dat,22)>> 7) & 0x7f;
//			}
			$tmp['xf']=$xf;
			if ($xf['formindex']==0) $tmp['format']='';
			else $tmp['format']=$this->recFORMAT[$xf['formindex']];

			$dat=$this->recFONT[$xf['fontindex']];
			$font['height']=$this->__get2($dat,4);
			$font['style']=$this->__get2($dat,6);
			$font['color']= $this->__get2($dat,8);
			$font['weight']=$this->__get2($dat,10);
			$font['escapement']=$this->__get2($dat,12);
			$font['underline']=$this->__get1($dat,14);
			$font['family']=$this->__get1($dat,15);
			$font['charset']=$this->__get1($dat,16);
			$font['fontname']=$this->cnvstring(substr($dat,18),1);
			$tmp['font']=$font;

			return $tmp;
		} else return null;
	}


	/**
	* Get sheet-name
	* @param integer $sn sheet number
	* @return string sheetname
	* @access public
	* @example ./sample_ex1.php sample_ex1
	*/
	function getSheetName($sn){
		return $this->cnvstring($this->boundsheets[$sn]['name'],1);
	}


	/**
	* Get Header
	* @param integer $sn sheet number
	* @return string header
	* @access public
	* @example ./sample_ex1.php sample_ex1
	*/
	function getHeader($sn){
		return $this->cnvstring(substr($this->sheetbin[$sn]['header'], 4),2);
	}


	/**
	* Get Footer
	* @param integer $sn sheet number
	* @return string footer
	* @access public
	* @example ./sample_ex1.php sample_ex1
	*/
	function getFooter($sn){
		return $this->cnvstring(substr($this->sheetbin[$sn]['footer'], 4),2);
	}


	/**
	* Get Row Height
	* @param integer $sn sheet number
	* @param integer $row Row position
	* @return integer row-height
	* @access public
	* @example ./sample_ex1.php sample_ex1
	*/
	function getRowHeight($sn,$row){
		if (isset($this->rowblock[$sn][$row]['height'])){
			$ret=$this->rowblock[$sn][$row]['height'];
		} else {
			$ret=$this->defrowH[$sn];
		}
		return $ret;
	}

	/**
	* Get Column Width
	* @param integer $sn sheet number
	* @param integer $col Column position
	* @return integer column-width
	* @access public
	* @example ./sample_ex1.php sample_ex1
	*/
	function getColWidth($sn,$col){
		if (isset($this->colblock[$sn][$col]['width'])){
			$ret=$this->colblock[$sn][$col]['width'];
		} else {
			$ret=$this->defcolW[$sn] * 256 + 256;
		}
		return $ret;
	}


	/**
	* @access private
	*/
	function _setPrintInfo(){
		if (count($this->prntitle)>0)
		foreach($this->prntitle as $sheet => $val){
			unset($this->boundsheets[$sheet]['namerecord'][7]);
			$area='';
			$tmp['sheetindex']= $sheet+1;
			$tmp['menu2name']=pack("H*",'000000000007');
			$tmp['remain']='';
			if ($val['col1st']!==null)
				$area.="3bX0000ffff".bin2hex(pack("vv",$val['col1st'],$val['collast']));
			if ($val['row1st']!==null)
				$area.="3bX" . bin2hex(pack("vv",$val['row1st'],$val['rowlast'])) ."0000ff00";
			if ($val['col1st']!==null && $val['row1st']!==null){
				$tmp['formula']="291700".$area."10";
				$tmp['flags2notu']=pack("H*",'18002a00200000011a000000');
			} else {
				$tmp['formula']=$area;
				$tmp['flags2notu']=pack("H*",'18001b00200000010b000000');
			}
			$this->boundsheets[$sheet]['namerecord'][7]=$tmp;
		}

		if (count($this->prnarea)<1) return;
		foreach($this->prnarea as $sheet => $val){
			unset($this->boundsheets[$sheet]['namerecord'][6]);
			$area='';
			$tmp['flags2notu']=pack("H*",'18001b00200000010b000000');
			$tmp['sheetindex']= $sheet+1;
			$tmp['menu2name']=pack("H*",'000000000006');
			$tmp['remain']='';
			$area.="3bX".bin2hex(pack("vvvv",$val['row1st'],$val['rowlast'],$val['col1st'],$val['collast']));
			$tmp['formula']=$area;
			$this->boundsheets[$sheet]['namerecord'][6]=$tmp;
		}
	}

	/**
	* @access private
	*/
	function _makesupblock(){
		if (count($this->dupsheet)+count($this->rmsheets)+count($this->prntitle)+count($this->prnarea) >0){
			$curnum=count($this->boundsheets);
			$this->globaldat['supbook']=pack("vvvv",Type_SUPBOOK,4,$curnum,0x401);
			$exsheetdat='';
			for($i=0;$i<$curnum;$i++){
				$exsheetdat.=pack("vvv",0,$i,$i);
			}
			$this->globaldat['extsheet']=pack("vvv",0x17,strlen($exsheetdat)+2,$curnum).$exsheetdat;
			$nr='';
			foreach((array)$this->boundsheets as $sn =>$tmp){
				if (isset($tmp['namerecord'][6]))
				if (count($tmp['namerecord'][6])>0){
					$nr.=$tmp['namerecord'][6]['flags2notu'].pack("v",$sn+1).$tmp['namerecord'][6]['menu2name'];
					$nr.=pack("H*",str_replace('X',bin2hex(pack('v',$sn)),$tmp['namerecord'][6]['formula']));
					$nr.=$tmp['namerecord'][6]['remain'];
				}
				if (isset($tmp['namerecord'][7]))
				if (count($tmp['namerecord'][7])>0){
					$nr.=$tmp['namerecord'][7]['flags2notu'].pack("v",$sn+1).$tmp['namerecord'][7]['menu2name'];
					$nr.=pack("H*",str_replace('X',bin2hex(pack('v',$sn)),$tmp['namerecord'][7]['formula']));
					$nr.=$tmp['namerecord'][7]['remain'];
				}
			}
			$this->globaldat['namerecord']=$nr;
		}
		return;
	}


	/**
	* @access private
	*/
	function analizeform($form){
		$fpos=0;
		$flen=strlen($form);
		$ret='';
		while ($fpos < $flen){
			$token=$this->__get1($form,$fpos);
			if ($token > 0x3F) $token -=0x20;
			if ($token > 0x3F) $token -=0x20;
			switch ($token){
			case 0x3:
			case 0x4:
			case 0x5:
			case 0x6:
			case 0x7:
			case 0x8:
			case 0x9:
			case 0xA:
			case 0xB:
			case 0xC:
			case 0xD:
			case 0xE:
			case 0xF:
			case 0x10:
			case 0x11:
			case 0x12:
			case 0x13:
			case 0x14:
			case 0x15:
			case 0x16:
				$ret.=bin2hex(substr($form,$fpos,1));
				$fpos+=1;
				break;
	//		case 0x17:
	//		case 0x18:
	//		case 0x19:
	//			$fpos = $flen;
	//			break;
			case 0x1C:
			case 0x1D:
				$ret.=bin2hex(substr($form,$fpos,2));
				$fpos+=2;
				break;
			case 0x1E:
			case 0x29:
			case 0x2E:
			case 0x2F:
			case 0x3D:
				$ret.=bin2hex(substr($form,$fpos,3));
				$fpos+=3;
				break;
			case 0x21:
				$ret.=bin2hex(substr($form,$fpos,4));
				$fpos+=4;
				break;
			case 0x1:
			case 0x2:
			case 0x22:
			case 0x23:
			case 0x24:
			case 0x2A:
			case 0x2C:
				$ret.=bin2hex(substr($form,$fpos,5));
				$fpos+=5;
				break;
			case 0x39:
			case 0x3A:
			case 0x3C:
				$ret.=bin2hex(substr($form,$fpos,1));
				$ret.="X";
				$ret.=bin2hex(substr($form,$fpos+3,4));
				$fpos+=7;
				break;
			case 0x26:
			case 0x27:
			case 0x28:
				$ret.=bin2hex(substr($form,$fpos,7));
				$fpos+=7;
				break;
			case 0x1F:
			case 0x20:
			case 0x25:
			case 0x2B:
			case 0x2D:
				$ret.=bin2hex(substr($form,$fpos,9));
				$fpos+=9;
				break;
			case 0x3B:
			case 0x3D:
				$ret.=bin2hex(substr($form,$fpos,1));
				$ret.="X";
				$ret.=bin2hex(substr($form,$fpos+3,8));
				$fpos+=11;
				break;
			default:
				$ret=bin2hex($form);
				$fpos = $flen;
			}
		}
		return $ret;
	}

	/**
	* Add Image to Sheet
	* @param integer $sn sheet number
	* @param integer $row Row position
	* @param integer $col Column posion  0indexed
	* @param string  $image path to the image file
	* @param integer $x horizontal offset pixel(option)
	* @param integer $y vertical offset pixel(option)
    * @param integer $scale_x The horizontal scale
    * @param integer $scale_y The vertical scale
	* @access public
	*/
    function addImage($sn, $row, $col, $image, $x = 0, $y = 0, $scale_x = 1, $scale_y = 1){
		$val['sheet']=$sn;
		$val['row']=$row;
		$val['col']=$col;
		$val['image']=$image;
		$val['dx']=$x;
		$val['dy']=$y;
		$val['scaleX']=$scale_x;
		$val['scaleY']=$scale_y;
		$this->revise_dat['add_image'][$sn][]=$val;
	}

    /**
    * @access private
    */
    function _posImage($sn,$colstart, $rowstart, $x1, $y1, $width, $height) {
        $colend = $colstart;
        $rowend = $rowstart;
        if ($x1 >= $this->_sizeCol($sn,$colstart)) $x1 = 0;
        if ($y1 >= $this->_sizeRow($sn,$rowstart)) $y1 = 0;
        $width += $x1;
        $height += $y1;
        while ($width >= $this->_sizeCol($sn,$colend))
            $width -= $this->_sizeCol($sn,$colend++);
        while ($height >= $this->_sizeRow($sn,$rowend))
            $height -= $this->_sizeRow($sn,$rowend++);
        if ($this->_sizeCol($sn,$colstart) == 0) return;
        if ($this->_sizeCol($sn,$colend) == 0) return;
        if ($this->_sizeRow($sn,$rowstart) == 0) return;
        if ($this->_sizeRow($sn,$rowend) == 0) return;
        $x1 = $x1 / $this->_sizeCol($sn,$colstart) * 1024;
        $y1 = $y1 / $this->_sizeRow($sn,$rowstart) *  256;
        $x2 = $width / $this->_sizeCol($sn,$colend) * 1024;
        $y2 = $height / $this->_sizeRow($sn,$rowend) *  256;
        $data  = pack("vvVvvv", 0x5d, 0x3c, 0x01, 0x08, 0x01, 0x614);
        $data .= pack("vvvv", $colstart, $x1, $rowstart, $y1);
        $data .= pack("vvvv", $colend, $x2, $rowend, $y2);
        $data .= pack("vVv", 0, 0, 0);
        $data .= pack("CCCCCCCC", 9, 9, 0, 0, 8, 0xff, 1, 0);
        $data .= pack("vVvvvvV", 0, 9, 0, 0, 0, 1, 0);
        return($data);
    }

    /**
    * @access private
    */
    function _getcolxf($sn,$col) {
		if (isset($this->colblock[$sn][$col])){
			$cxf = $this->colblock[$sn][$col]['xf'];
		} else {
			$cxf = 0x0f;
		}
		return $cxf;
	}

    /**
    * @access private
    */
    function _sizeCol($sn,$col) {
		$c=$this->getColWidth($sn,$col);
		return ($c == 0) ? 0 : (floor(8 * $c / 256 + 0));
	}

    /**
    * @access private
    */
    function _sizeRow($sn,$row) {
		$r=$this->getRowHeight($sn,$row);
		return ($r == 0) ? 0 :(floor($r / 15));
	}

    /**
    * @access private
    */
    function _makeImageOBJ($sn) {
		$tmp="";
		if (isset($this->revise_dat['add_image'][$sn]))
		foreach($this->revise_dat['add_image'][$sn] as $val)
			$tmp.=$this->_addImage($val['sheet'],$val['row'],$val['col'], $val['image'],
			 $val['dx'], $val['dy'], $val['scaleX'], $val['scaleY']);
		return($tmp);
	}

    /**
    * @access private
    */
    function _addImage($sn, $row, $col, $imgname, $x1=0, $y1=0, $scale_x=1, $scale_y=1){
		$ermes=array();
		$ext=strtolower(substr($imgname,-4));
		$im='';
		switch($ext){
			case ".jpg":
			case "jpeg":
				if (function_exists('imagecreatefromjpeg'))
					$im = @imagecreatefromjpeg($imgname);
				break;
			case ".gif":
				if (function_exists('imagecreatefromgif'))
					$GIFim = @imagecreatefromgif($imgname);
					$im = imagecreatetruecolor(imagesx ($GIFim), imagesy ($GIFim));
					imagecopy($im, $GIFim, 0, 0, 0, 0, imagesx ($GIFim), imagesy ($GIFim));
					imagedestroy($GIFim);
				break;
			case ".png":
				if (function_exists('imagecreatefrompng'))
					$im = @imagecreatefrompng($imgname);
				break;
			case ".bmp";
		        // Open file.
		        if (!$bmh = @fopen($imgname,"rb")) {
					$ermes[]="Couldn't open $imgname";
					break;
				}
				if ($this->Flag_Magic_Quotes) set_magic_quotes_runtime(0);
		        $data = fread($bmh, filesize($imgname));
				fclose($bmh);
				if ($this->Flag_Magic_Quotes) set_magic_quotes_runtime($this->Flag_Magic_Quotes);
		        if (strlen($data) <= 0x36) {
					$ermes[]="$imgname is too small.";
					break;
				}
		        if (substr($data,0,2) != "BM") $ermes[]="$imgname isn't BMP image.";
				$planes = $this->__get2($data,26);
				$bits  = $this->__get2($data,28);
		        $compress = $this->__get4($data,30);
		        if ($planes != 1) $ermes[]="$imgname: only 1 plane supported.";
		        if ($compress != 0) $ermes[]="$imgname: compression not supported.";
		        $size   = $this->__get4($data,2) - 0x36 + 0x0C;
		        $width  = $this->__get4($data,18);
		        $height = $this->__get4($data,22);
				if ( count($ermes)==0 )
		        if ($bits == 24 ) {
			        $data = substr($data, 0x36);
					break;
				} else {
					$ermes[]="$imgname isn't a 24bit color.";
					$data='';
					break;
				}
			default:
				$ermes[]="$imgname is unknown image type.";
		}
		if ($im){
			$width = imagesx ($im);
			$height = imagesy ($im);
			$BPLine = $width * 3;
			$Stride = ($BPLine + 3) & ~3;
			$size = $Stride * $height + 0x0C;
			$data='';
			$numpad = $Stride - $BPLine;
			for ($y = $height - 1; $y >= 0; --$y) {
				for ($x = 0; $x < $width; ++$x) {
					$colr = imagecolorat ($im, $x, $y);
					$data .= substr(pack ('V', $colr), 0,3);
				}
				for ($i = 0; $i < $numpad; ++$i)
					$data .= pack ('C', 0);
			}
		}
		if ($width <1 || $height<1 || strlen($data) < 3) $ermes[]="$imgname: Too small size";
		if ($width > 0xFFFF) $ermes[]="$imgname: Too large width";
		if ($height > 0xFFFF) $ermes[]="$imgname: Too large height";
		if (isset($ermes))
		if (count($ermes) > 0) {
			if ($this->debug_image){
				print_r($ermes);
				exit;
			}
			return '';
		}
		$data  = pack("Vvvvv", 0x000c, $width, $height, 0x01, 0x18) . $data;
        $width  *= $scale_x;
        $height *= $scale_y;

        $recOBJ=$this->_posImage($sn, $col, $row, $x1, $y1, $width, $height);

		$recCONT="";
		if ($size >0x814){
			$st=0x814;
			$header = pack("vvvvV", 0x007f, 8 + 0x814, 0x09, 0x01, $size);
			$recIMDATA = $header.substr($data,0,0x814);
			while($st+0x814 < $size){
				$recCONT .= pack("vv",0x3c,0x814).substr($data,$st,0x814);
				$st+=0x814;
			}
			if ($st< $size)
				$recCONT .= pack("vv",0x3c,$size-$st).substr($data,$st);
		} else {
			$header = pack("vvvvV", 0x007f, 8 + $size, 0x09, 0x01, $size);
			$recIMDATA = $header.$data;
		}
		unset($im);
		return $recOBJ.$recIMDATA.$recCONT;
    }

	/**
	* Set(Get) Error-Handling Method.
	* @param integer $mode error handling method(default 0)
	* @return integer error handling method
	* @access public
	*/
	function setErrorHandling($mode=''){
		if (is_numeric($mode)) {
			$this->Flag_Error_Handling = $mode;
		}
		return $this->Flag_Error_Handling;
	}

	/**
	* Set(Get) Flag_inherit_Info
	* @param integer $mode 1:inherit property (default 0)
	* @return integer flag-value
	* @access public
	*/
	function setInheritInfomation($mode=''){
		if (is_numeric($mode)) {
			$this->Flag_inherit_Info = $mode;
		}
		return $this->Flag_inherit_Info;
	}

	/**
	* @param mixed $data object
	* @return boolean True:The error occurred
	* @access public
	*/
    function isError($data){return is_a($data, 'ErrMess');}

	/**
	* @access private
	*/
    function &raiseError($message = ''){
		if ($this->Flag_Error_Handling == 0){
			die($message);
		}
		return new ErrMess($message);
	}

	/**
	* @param mixed $dat ole-stream
	* @return array property-data
	* @access public
	*/
	function getProperty($dat){
		$tmp=array();
		$tmp['header']=substr($dat, 0, 24);
$tmp['header']=bin2hex(substr($dat, 0, 24));
		$secs =  $this->__get4($dat, 0x18);
		if ($secs < 1) return;
		$pos=0x1C;
		$sec=array();
		for ($i = 0; $i < $secs; $i++){
			$sec['FMTID']= bin2hex(substr($dat, $pos, 16));
			$sec['offset']= $this->__get4($dat, $pos + 16);
			$tmp['section'][$i]=$sec;
			$pos += 20;
		}
		for ($i = 0; $i < $secs; $i++){
			$spos=$tmp['section'][$i]['offset'];
			$Props = $this->__get4($dat, $spos + 4);
			$spos +=8;
			for ($j = 0; $j < $Props; $j++){
				$prop['propid']=$this->__get4($dat, $spos);
				$ppos = $this->__get4($dat, $spos + 4) + $tmp['section'][$i]['offset'];
				$spos +=8;
				$prop['Type']=$this->__get4($dat, $ppos);
				if ($prop['propid']==0){ //dictionary
					$dpos=0;
					$numdic=$prop['Type'];
//print "numdic=". $numdic ."\n";
					for ($k=0; $k<$numdic; $k++){
						$dlen=$this->__get4($dat, $ppos+$dpos+8);
						$cusname=mb_convert_encoding(substr($dat,$ppos+$dpos+12,$dlen), $this->charset,"SJIS-win");
$cusname=substr($dat,$ppos+$dpos+12,$dlen);
						$dprop[$this->__get4($dat, $ppos+$dpos+4)]=trim($cusname);
						$dpos+=$dlen+8;
					}
//					$tmp['section'][$i]['dic']=$dprop;
					$prop['dic']=$dprop;
					unset($dprop);
				} else {
					switch ($prop['Type']) {
					case 2:	// si ds
						$prop['Rgb']=substr($dat,$ppos+4,4);
						break;
					case 3:	// si ds
						$prop['Rgb']=substr($dat,$ppos+4,4);
						break;
					case 6:	//
						$prop['Rgb']='';
						break;
					case 11:	// ds
						$prop['Rgb']=substr($dat,$ppos+4,4);
						break;
					case 12:	// ds
						$prop['Rgb']=substr($dat,$ppos+4,4);
						break;

					case 30: // si
						$len=$this->__get4($dat, $ppos+4);
						$prop['Rgb']=substr($dat,$ppos+4,$len+4);
						break;
					case 64:	// si
						$prop['Rgb']=substr($dat,$ppos+4,8);
						break;
					case 65:	// ds
						$len=$this->__get4($dat, $ppos+4);
						$prop['Rgb']=substr($dat,$ppos+4,$len+4);
						break;
					case 71:	// si
						$prop['Rgb']='';
						break;
					case 4126:	// ds
						$parts=array();
						$n=$this->__get4($dat, $ppos+4);
						$p=$ppos+8;
						for ($k=0;$k<$n;$k++){
							$parts[$k]=substr($dat,$p+4,$this->__get4($dat, $p));
							$p+=4+$this->__get4($dat, $p);
						}
						$prop['Parts']=$parts;
						$prop['Rgb']=substr($dat,$ppos+4,$p-$ppos-4);
						break;
					case 4108:	// ds
						$s ='02000000';
						$s.='1E0000000D000000838F815B834E8356815B836700';
						$s.='0300000002000000';
						$prop['Rgb']=pack("H*",$s);
						break;
					default:
					}
				}
$prop['Rgb']=bin2hex($prop['Rgb']);
				$tmp['section'][$i]['prop'][$j]=$prop;
				unset($prop);
			}
		}
//print_r($tmp);
		return $tmp;
	}

/**
* @access private
*/
	function makehash16b($password){
// The length of the password is restricted to 15 characters.
//		$password ="reviser";
		$char_index=strlen($password)-1;
		if ($char_index < 1 || $char_index >15) return;
		$hash=0;
		for( $i=$char_index; $i>=0; $i-- ){
			$char=substr($password,$i,1);
			$hash ^=ord($char);
			$hash *=2;
			if($hash & 0x8000) {
				$hash &= 0x7fff;
				$hash++;
			}
		}
		$hash ^= 0xCE4B ^ strlen($password);
//		 echo "hash=".dechex($hash);
		return $hash;
	}

/**
* @access private
*/
	function decrypt(&$dat,&$datx){
		$pos = 0;
		$code=-1;
		$tmp='';
		$poslimit=strlen($dat);
		while ($code != 0){
			if ($pos > $poslimit){
				break;
			}
		    $code = $this->__get2($dat,$pos);
		    $length = $this->__get2($dat,$pos+2);
		    switch ($code) {
			case Type_BOF:
			case Type_INTERFACEHDR:
			case Type_FILEPASS:
			case Type_BOF:
				$tmp.=substr($dat, $pos, $length+4);
			    break;
			case Type_BOUNDSHEET:
				$tmp.=substr($dat, $pos, 8);
				$tmp.=substr($datx, $pos+8, $length-4);
			    break;
			default:
				$tmp.=substr($dat, $pos, 4);
				$tmp.=substr($datx, $pos+4, $length);
			}
			$pos += $length + 4;
		}
		return $tmp;
	}

/**
* @access private
*/
	function encrypt(&$dat){
		if (strlen($this->saveFPass)<1) return $this->raiseError("No Password");
		$pos = 0;
		$code = -1;
		$length = $this->__get2($dat,$pos+2);
		$tmp=substr($dat, $pos, $length+4);
		$pos += $length + 4;
		$docid=md5(date("ymdHis")); // docid
		$salt="36832c6f5130c930c330c830b330e030"; // salt
		$docid=pack("H*",$docid);
		$salt=pack("H*",$salt);
		$md5= NEW XL_Crypto();
		$hashedsalt=$md5->makehashedsalt($this->saveFPass, $docid, $salt);
		if (strlen($hashedsalt) != 0x10) return $this->raiseError("Unknown ERROR! hash-make");
		$tmp.=pack("H*","2f003600010001000100").$docid.$salt.$hashedsalt;
		$size=strval(strlen($dat) + 0x3A);
		$op = $md5->makeXorPtn($size);
		if (strlen($op) < $size) return $this->raiseError("XOR-Index make-error");
		$datx=$dat ^ (substr($op,0x3a));
		unset($op);
		$poslimit=strlen($dat)+0x3a;
		while ($code != 0){
			if ($pos > $poslimit){
				break;
			}
		    $code = $this->__get2($dat,$pos);
		    $length = $this->__get2($dat,$pos+2);
		    switch ($code) {
			case Type_BOF:
			case Type_INTERFACEHDR:
			case Type_FILEPASS:
			case Type_BOF:
				$tmp.=substr($dat, $pos, $length+4);
			    break;
			case Type_BOUNDSHEET:
				$soffset = $this->__get4($dat, $pos+4);
				$tmp.=substr($dat, $pos, 4);
				$tmp.=pack("V",$soffset+0x3a);
				$tmp.=substr($datx, $pos+8, $length-4);
			    break;
			default:
				$tmp.=substr($dat, $pos, 4);
				$tmp.=substr($datx, $pos+4, $length);
			}
			$pos += $length + 4;
		}
		unset($datx);
		$dat=$tmp;
		return;
	}
}

/**
* error class
*/
class ErrMess {
    var $message = '';
	/**
	* @param string $message Error message
	* @access public
	*/
    function ErrMess($message){$this->message = $message;}
	/**
	* @return string Error message
	* @access public
	*/
    function getMessage() {return ($this->message);}
}

/**
* MD5 work structure for Encryption/Encryption Class
*/
class MD5_CTX
{
	var $i=array(0,0);
	var $buf=array(0,0,0,0);
	var $in='';
	var $digest='';

	function MD5_CTX(){
		$this->in=str_repeat("\x00",64);
		$this->digest=str_repeat("\x00",16);
	}
}

/**
* RC4 work structure for Encryption/Encryption Class
*/
class rc4_key
{
	var $x;
	var $y;
	var $state='';

	function rc4_key(){
		$this->state=str_repeat("\x00",256);
	}
}

/**
* Encryption/Encryption Class
* MD5 message-digest algorithm
*/
class XL_Crypto
{
	var $PADDING=array();
	var $valContext;

	function XL_Crypto()
	{
		$this->PADDING=str_repeat("\x00",64);
		$this->PADDING[0]=chr(0x80);
		$this->valContext= NEW MD5_CTX;
	}

	/**
	* general function
	* @access private
	*/
	function memcpy(&$dest,$offset,$org,$org_offset,$n){
		for ($i=0;$i<$n;$i++){
			$dest[$offset+$i]=$org[$org_offset+$i];
		}
	}

	/**
	* @access private
	*/
	function memset(&$dest,$offset,$char,$n){
		for ($i=0;$i<$n;$i++){
			$dest[$offset+$i]=$char;
		}
	}

	/**
	* Start MD5 accumulation.  Set bit count to 0 and buffer to mysterious
	* initialization constants.
	* @access private
	*/
	function MD5Init(&$mdContext)
	{
		$mdContext->buf[0] = 0x067452301;
		$mdContext->buf[1] = 0x0efcdab89;
		$mdContext->buf[2] = 0x098badcfe;
		$mdContext->buf[3] = 0x010325476;

		$mdContext->i[0] = 0;
		$mdContext->i[1] = 0;
	}

	/**
	* @access private
	*/
	function MD5Update(&$mdContext, $inBuf, $inLen)
	{
  /* compute number of bytes mod 64 */
		$mdi = ($mdContext->i[0] >> 3) & 0x3f;
  /* update number of bits */
		if (($mdContext->i[0] + ($inLen << 3)) < $mdContext->i[0]) $mdContext->i[1]++;

		$mdContext->i[0] += ($inLen << 3);
		$mdContext->i[1] += ($inLen >> 29) & 0x07;
		$pos=0;
		while ($inLen--) {
    /* add new character to buffer, increment mdi */
			$mdContext->in[$mdi++] = $inBuf[$pos++];
    /* transform if necessary */
			if ($mdi == 0x40){
				$ii = 0;
				for ($i=0;$i < 16;$i++){
					$in[$i] = (ord($mdContext->in[$ii+3]) << 24) |
						(ord($mdContext->in[$ii+2]) << 16) |
						(ord($mdContext->in[$ii+1]) << 8) |
						 ord($mdContext->in[$ii]);
					$ii +=4;
				}

				$this->Transform($mdContext->buf,$in);
				$mdi=0;
			}
		}
	}

	/**
	* @access private
	*/
	function MD5Final (&$mdContext)	// This function is not complete yet
	{
	  $in=array(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);

	  /* save number of bits */
	  $in[14] = $mdContext->i[0];
	  $in[15] = $mdContext->i[1];

	  /* compute number of bytes mod 64 */
	  $mdi = (int)(($mdContext->i[0] >> 3) & 0x3F);

	  /* pad out to 56 mod 64 */
	  $padLen = ($mdi < 56) ? (56 - $mdi) : (120 - $mdi);
	  $this->MD5Update($mdContext, $this->PADDING, $padLen);

	  /* append length in bits and transform */
	  for ($i = 0; $i < 14; $i++){
	    $in[$i] = (ord($mdContext->in[$i*4+3]) << 24) |
	            (ord($mdContext->in[$i*4+2]) << 16) |
	            (ord($mdContext->in[$i*4+1]) << 8) |
	            ord($mdContext->in[$i*4]);
	  }
	  $this->Transform ($mdContext->buf, $in);

	  /* store buffer in digest */
	  for ($i = 0; $i < 4; $i++) {
	    $mdContext->digest[$i*4] = chr($mdContext->buf[$i] & 0xFF);
	    $mdContext->digest[$i*4+1] =chr(($mdContext->buf[$i] >> 8) & 0xFF);
	    $mdContext->digest[$i*4+2] =chr(($mdContext->buf[$i] >> 16) & 0xFF);
	    $mdContext->digest[$i*4+3] =chr(($mdContext->buf[$i] >> 24) & 0xFF);
	  }
	}

	/**
	* @access private
	*/
	function F($x,$y,$z) {
		$t[0]=($x[0] & $y[0]) | ((0xffff ^ $x[0]) & $z[0]);
		$t[1]=($x[1] & $y[1]) | ((0xffff ^ $x[1]) & $z[1]);
		return $t;
	}
	/**
	* @access private
	*/
	function G($x,$y,$z) {
		$t[0]=($x[0] & $z[0]) | ($y[0] & (0xffff ^ $z[0]));
		$t[1]=($x[1] & $z[1]) | ($y[1] & (0xffff ^ $z[1]));
		return $t;
	}
	/**
	* @access private
	*/
	function H($x,$y,$z) {
		$t[0]=($x[0] ^ $y[0] ^ $z[0]);
		$t[1]=($x[1] ^ $y[1] ^ $z[1]);
		return $t;
	}
	/**
	* @access private
	*/
	function I($x,$y,$z) {
		$t[0]=$y[0] ^ ($x[0] | (0xffff ^ $z[0]));
		$t[1]=$y[1] ^ ($x[1] | (0xffff ^ $z[1]));
		return $t;
	}

	/**
	* @access private
	*/
	function ROTATE_LEFT($x, $n) {
		if ($n==0 | $n==32) return $x;
		if ($n==16) return array($x[1],$x[0]);
		if ($n>=16) {
			$n -=16;
			$f =0; } else $f=1;
		$b[0]=($x[0] << $n);
		$b[1]=($x[1] << $n);
		$t[0]=($x[0] & 0xffff) >> (16-$n);
		$t[1]=($x[1] & 0xffff) >> (16-$n);
		$a[0]=($b[0] | $t[1]) & 0xffff;
		$a[1]=($b[1] | $t[0]) & 0xffff;
		if ($f) return $a;
		return array($a[1],$a[0]);
	}

	/**
	* @access private
	*/
	function s2($a,$b){
		$a[0]+=$b[0];
		$a[1]+=$b[1];
		if ($a[0]>0xffff){ $a[1]++;$a[0]&=0xffff;}
		$a[1]&=0xffff;
		return $a;
	}

	/**
	* @access private
	*/
	function s4($a,$b,$c,$d){
		$a[0]+=$b[0]+$c[0]+$d[0];
		if ($a[0]>0xffff){ $a[1]+=($a[0]>>16)&0xffff;$a[0]&=0xffff;}
		$a[1]+=$b[1]+$c[1]+$d[1];
		$a[1]&=0xffff;
		return $a;
	}

	/**
	* @access private
	*/
	function ar($x){
		$a[0]=$x & 0xffff;
		$a[1]=($x >> 16) & 0xffff;
		return $a;
	}

	/**
	* @access private
	*/
	function FF(&$a, $b, $c, $d, $x, $s, $ac) {
		$a = $this->s4($a, $this->F($b, $c, $d), $this->ar($x) , $this->ar((int)$ac));
		$a = $this->ROTATE_LEFT($a, $s);
		$a = $this->s2($a,$b);
	  }

	/**
	* @access private
	*/
	function GG(&$a, $b, $c, $d, $x, $s, $ac){
		$a = $this->s4($a, $this->G($b, $c, $d), $this->ar($x) , $this->ar((int)$ac));
		$a = $this->ROTATE_LEFT($a, $s);
		$a = $this->s2($a,$b);
	  }

	/**
	* @access private
	*/
	function HH(&$a, $b, $c, $d, $x, $s, $ac){
		$a = $this->s4($a, $this->H($b, $c, $d), $this->ar($x) , $this->ar((int)$ac));
		$a = $this->ROTATE_LEFT($a, $s);
		$a = $this->s2($a,$b);
	  }

	/**
	* @access private
	*/
	function II(&$a, $b, $c, $d, $x, $s, $ac){
		$a = $this->s4($a, $this->I($b, $c, $d), $this->ar($x) , $this->ar((int)$ac));
		$a = $this->ROTATE_LEFT($a, $s);
		$a = $this->s2($a,$b);
	  }

	/**
	* @access private
	*/
	function Transform (&$buf, &$in)
	{
		$a0 = $a = $this->ar($buf[0]);
		$b0 = $b = $this->ar($buf[1]);
		$c0 = $c = $this->ar($buf[2]);
		$d0 = $d = $this->ar($buf[3]);

	  /* Round 1 */
	  $this->FF ( $a, $b, $c, $d, $in[ 0],  7, 0x0d76aa478); /* 1 */
	  $this->FF ( $d, $a, $b, $c, $in[ 1], 12, 0x0e8c7b756); /* 2 */
	  $this->FF ( $c, $d, $a, $b, $in[ 2], 17, 0x0242070db); /* 3 */
	  $this->FF ( $b, $c, $d, $a, $in[ 3], 22, 0x0c1bdceee); /* 4 */
	  $this->FF ( $a, $b, $c, $d, $in[ 4],  7, 0x0f57c0faf); /* 5 */
	  $this->FF ( $d, $a, $b, $c, $in[ 5], 12, 0x04787c62a); /* 6 */
	  $this->FF ( $c, $d, $a, $b, $in[ 6], 17, 0x0a8304613); /* 7 */
	  $this->FF ( $b, $c, $d, $a, $in[ 7], 22, 0x0fd469501); /* 8 */
	  $this->FF ( $a, $b, $c, $d, $in[ 8],  7, 0x0698098d8); /* 9 */
	  $this->FF ( $d, $a, $b, $c, $in[ 9], 12, 0x08b44f7af); /* 10 */
	  $this->FF ( $c, $d, $a, $b, $in[10], 17, 0x0ffff5bb1); /* 11 */
	  $this->FF ( $b, $c, $d, $a, $in[11], 22, 0x0895cd7be); /* 12 */
	  $this->FF ( $a, $b, $c, $d, $in[12],  7, 0x06b901122); /* 13 */
	  $this->FF ( $d, $a, $b, $c, $in[13], 12, 0x0fd987193); /* 14 */
	  $this->FF ( $c, $d, $a, $b, $in[14], 17, 0x0a679438e); /* 15 */
	  $this->FF ( $b, $c, $d, $a, $in[15], 22, 0x049b40821); /* 16 */

	  /* Round 2 */
	  $this->GG ( $a, $b, $c, $d, $in[ 1],  5, 0x0f61e2562); /* 17 */
	  $this->GG ( $d, $a, $b, $c, $in[ 6],  9, 0x0c040b340); /* 18 */
	  $this->GG ( $c, $d, $a, $b, $in[11], 14, 0x0265e5a51); /* 19 */
	  $this->GG ( $b, $c, $d, $a, $in[ 0], 20, 0x0e9b6c7aa); /* 20 */
	  $this->GG ( $a, $b, $c, $d, $in[ 5],  5, 0x0d62f105d); /* 21 */
	  $this->GG ( $d, $a, $b, $c, $in[10],  9, 0x002441453); /* 22 */
	  $this->GG ( $c, $d, $a, $b, $in[15], 14, 0x0d8a1e681); /* 23 */
	  $this->GG ( $b, $c, $d, $a, $in[ 4], 20, 0x0e7d3fbc8); /* 24 */
	  $this->GG ( $a, $b, $c, $d, $in[ 9],  5, 0x021e1cde6); /* 25 */
	  $this->GG ( $d, $a, $b, $c, $in[14],  9, 0x0c33707d6); /* 26 */
	  $this->GG ( $c, $d, $a, $b, $in[ 3], 14, 0x0f4d50d87); /* 27 */
	  $this->GG ( $b, $c, $d, $a, $in[ 8], 20, 0x0455a14ed); /* 28 */
	  $this->GG ( $a, $b, $c, $d, $in[13],  5, 0x0a9e3e905); /* 29 */
	  $this->GG ( $d, $a, $b, $c, $in[ 2],  9, 0x0fcefa3f8); /* 30 */
	  $this->GG ( $c, $d, $a, $b, $in[ 7], 14, 0x0676f02d9); /* 31 */
	  $this->GG ( $b, $c, $d, $a, $in[12], 20, 0x08d2a4c8a); /* 32 */

	  /* Round 3 */
	  $this->HH ( $a, $b, $c, $d, $in[ 5],  4, 0x0fffa3942); /* 33 */
	  $this->HH ( $d, $a, $b, $c, $in[ 8], 11, 0x08771f681); /* 34 */
	  $this->HH ( $c, $d, $a, $b, $in[11], 16, 0x06d9d6122); /* 35 */
	  $this->HH ( $b, $c, $d, $a, $in[14], 23, 0x0fde5380c); /* 36 */
	  $this->HH ( $a, $b, $c, $d, $in[ 1],  4, 0x0a4beea44); /* 37 */
	  $this->HH ( $d, $a, $b, $c, $in[ 4], 11, 0x04bdecfa9); /* 38 */
	  $this->HH ( $c, $d, $a, $b, $in[ 7], 16, 0x0f6bb4b60); /* 39 */
	  $this->HH ( $b, $c, $d, $a, $in[10], 23, 0x0bebfbc70); /* 40 */
	  $this->HH ( $a, $b, $c, $d, $in[13],  4, 0x0289b7ec6); /* 41 */
	  $this->HH ( $d, $a, $b, $c, $in[ 0], 11, 0x0eaa127fa); /* 42 */
	  $this->HH ( $c, $d, $a, $b, $in[ 3], 16, 0x0d4ef3085); /* 43 */
	  $this->HH ( $b, $c, $d, $a, $in[ 6], 23, 0x004881d05); /* 44 */
	  $this->HH ( $a, $b, $c, $d, $in[ 9],  4, 0x0d9d4d039); /* 45 */
	  $this->HH ( $d, $a, $b, $c, $in[12], 11, 0x0e6db99e5); /* 46 */
	  $this->HH ( $c, $d, $a, $b, $in[15], 16, 0x01fa27cf8); /* 47 */
	  $this->HH ( $b, $c, $d, $a, $in[ 2], 23, 0x0c4ac5665); /* 48 */

	  /* Round 4 */
	  $this->II ( $a, $b, $c, $d, $in[ 0],  6, 0x0f4292244); /* 49 */
	  $this->II ( $d, $a, $b, $c, $in[ 7], 10, 0x0432aff97); /* 50 */
	  $this->II ( $c, $d, $a, $b, $in[14], 15, 0x0ab9423a7); /* 51 */
	  $this->II ( $b, $c, $d, $a, $in[ 5], 21, 0x0fc93a039); /* 52 */
	  $this->II ( $a, $b, $c, $d, $in[12],  6, 0x0655b59c3); /* 53 */
	  $this->II ( $d, $a, $b, $c, $in[ 3], 10, 0x08f0ccc92); /* 54 */
	  $this->II ( $c, $d, $a, $b, $in[10], 15, 0x0ffeff47d); /* 55 */
	  $this->II ( $b, $c, $d, $a, $in[ 1], 21, 0x085845dd1); /* 56 */
	  $this->II ( $a, $b, $c, $d, $in[ 8],  6, 0x06fa87e4f); /* 57 */
	  $this->II ( $d, $a, $b, $c, $in[15], 10, 0x0fe2ce6e0); /* 58 */
	  $this->II ( $c, $d, $a, $b, $in[ 6], 15, 0x0a3014314); /* 59 */
	  $this->II ( $b, $c, $d, $a, $in[13], 21, 0x04e0811a1); /* 60 */
	  $this->II ( $a, $b, $c, $d, $in[ 4],  6, 0x0f7537e82); /* 61 */
	  $this->II ( $d, $a, $b, $c, $in[11], 10, 0x0bd3af235); /* 62 */
	  $this->II ( $c, $d, $a, $b, $in[ 2], 15, 0x02ad7d2bb); /* 63 */
	  $this->II ( $b, $c, $d, $a, $in[ 9], 21, 0x0eb86d391); /* 64 */

		$a =$this->s2($a,$a0);
		$b =$this->s2($b,$b0);
		$c =$this->s2($c,$c0);
		$d =$this->s2($d,$d0);

	  $buf[0] = ($a[1]<<16)|$a[0];
	  $buf[1] = ($b[1]<<16)|$b[0];
	  $buf[2] = ($c[1]<<16)|$c[0];
	  $buf[3] = ($d[1]<<16)|$d[0];
	}

	/**
	* @access private
	*/
	function MD5StoreDigest(&$mdContext)
	{
	/* store buffer in digest */
		$ii=0;
		for ($i = 0; $i < 4; $i++){
			$mdContext->digest[$ii] = chr($mdContext->buf[$i] & 0xFF);
			$mdContext->digest[$ii+1] =chr(($mdContext->buf[$i] >> 8) & 0xFF);
			$mdContext->digest[$ii+2] =chr(($mdContext->buf[$i] >> 16) & 0xFF);
			$mdContext->digest[$ii+3] =chr(($mdContext->buf[$i] >> 24) & 0xFF);
			$ii +=4;
		}
	}

	/**
	* @access private
	*/
function makekey($block, $key)
	{
	$mdContext=NEW MD5_CTX;
    $pwarray =str_repeat("\x00",64);

    /* 40 bit of hashed password, set by verifypwd() */
    $this->memcpy($pwarray,0, $this->valContext->digest,0, 5);

    /* put block number in byte 6...9 */
	$pwarray[5] = chr($block & 0xFF);         
	$pwarray[6] = chr(($block >>  8) & 0xFF); 
	$pwarray[7] = chr(($block >> 16) & 0xFF); 
	$pwarray[8] = chr(($block >> 24) & 0xFF); 

    $pwarray[9] = chr(0x80);
    $pwarray[56] = chr(0x48);

	$this->MD5Init ($mdContext);
	$this->MD5Update ($mdContext, $pwarray, 64);
	$this->MD5StoreDigest($mdContext);
    $this->prepare_key($mdContext->digest, 16, $key);
	}

	/**
	* @access private
	*/
function prepare_key($key_data_ptr, $key_data_len, &$key)
{
   for($counter = 0; $counter < 256; $counter++){
   $key->state[$counter] = chr($counter);}
   $key->x = 0;
   $key->y = 0;
   $index1 = 0;
   $index2 = 0;
   for($counter = 0; $counter < 256; $counter++)
   {
      $index2 = (ord($key_data_ptr[$index1]) + ord($key->state[$counter]) + $index2) & 0xff;
		$x=ord($key->state[$counter]);
		$key->state[$counter]=$key->state[$index2];
		$key->state[$index2]=chr($x);
      $index1 = ($index1 + 1) % $key_data_len;  
   }       
}

	/**
	* @access private
	*/
function rc4(&$buffer_ptr, $buffer_len, &$key)
{ 
   $x = $key->x;
   $y = $key->y;
   for($counter = 0; $counter < $buffer_len; $counter++)
   {
      $x = ($x + 1) % 256;
      $y = (ord($key->state[$x]) + $y) % 256;
		$swp=$key->state[$x];
		$key->state[$x]=$key->state[$y];
		$key->state[$y]=$swp;
      $xorIndex = (ord($key->state[$x]) + ord($key->state[$y])) % 256;

      $buffer_ptr[$counter] = chr(ord($buffer_ptr[$counter])^ ord($key->state[$xorIndex]));
   }
   $key->x = $x;
   $key->y = $y;
}

	/**
	* @access private
	*/
function verifypwd($pwd, $docid, $salt, $hashedsalt)
	{
	$mdContext1=NEW MD5_CTX;
	$mdContext2=NEW MD5_CTX;
    $key = NEW rc4_key;
	$pwarray="";

	$pwd .= chr(0);
	$this->expandpwb($pwd,$pwarray);

	$this->MD5Init ($mdContext1);
	$this->MD5Update($mdContext1, $pwarray, 64);
	$this->MD5StoreDigest($mdContext1);

    $offset = 0;
    $keyoffset = 0;
    $tocopy = 5;

	$this->MD5Init($this->valContext);

    while ($offset != 16)
    {
        if ((64 - $offset) < 5)
            $tocopy = 64 - $offset;

        $this->memcpy ($pwarray, $offset, $mdContext1->digest, $keyoffset, $tocopy);
        $offset += $tocopy;

        if ($offset == 64)
        {
			$this->MD5Update ($this->valContext, $pwarray, 64);
            $keyoffset = $tocopy;
            $tocopy = 5 - $tocopy;
            $offset = 0;
            continue;
        }

        $keyoffset = 0;
        $tocopy = 5;
        $this->memcpy ($pwarray, $offset, $docid, 0, 16);
        $offset += 16;
    }

    /* Fix (zero) all but first 16 bytes */

    $pwarray[16] = chr(0x80);
    $this->memset($pwarray, 17, chr(0), 47);
    $pwarray[56] = chr(0x80);
    $pwarray[57] = chr(0x0A);

	$this->MD5Update ($this->valContext, $pwarray, 64);
	$this->MD5StoreDigest($this->valContext);

    /* Generate 40-bit RC4 key from 128-bit hashed password */

	$this->makekey(0, $key);
    $this->rc4 ($salt, 16, $key);

    $this->rc4 ($hashedsalt, 16, $key);

    $salt[16] = chr(0x80);
    $this->memset($salt, 17, chr(0), 47);
    $salt[56] = chr(0x80);

	$this->MD5Init ($mdContext2);
	$this->MD5Update ($mdContext2, $salt, 64);
	$this->MD5StoreDigest($mdContext2);
    return (($mdContext2->digest==$hashedsalt)? -1:0);
	}

	/**
	* @access private
	*/
	function makehashedsalt($pwd, $docid, $salt){
		$mdContext1=NEW MD5_CTX;
		$mdContext2=NEW MD5_CTX;
		$key = NEW rc4_key;
		$hashedsalt=str_repeat("\x00",16);
		$pwarray="";

		$pwd .= chr(0);
		$this->expandpwb($pwd,$pwarray);
		$this->MD5Init ($mdContext1);
		$this->MD5Update($mdContext1, $pwarray, 64);
		$this->MD5StoreDigest($mdContext1);
		$offset = 0;
		$keyoffset = 0;
		$tocopy = 5;
		$this->MD5Init($this->valContext);
		while ($offset != 16){
			if ((64 - $offset) < 5) $tocopy = 64 - $offset;
			$this->memcpy ($pwarray, $offset, $mdContext1->digest, $keyoffset, $tocopy);
			$offset += $tocopy;
			if ($offset == 64){
				$this->MD5Update ($this->valContext, $pwarray, 64);
				$keyoffset = $tocopy;
				$tocopy = 5 - $tocopy;
				$offset = 0;
				continue;
			}
			$keyoffset = 0;
			$tocopy = 5;
			$this->memcpy ($pwarray, $offset, $docid, 0, 16);
			$offset += 16;
		}
		$pwarray[16] = chr(0x80);
		$this->memset($pwarray, 17, chr(0), 47);
		$pwarray[56] = chr(0x80);
		$pwarray[57] = chr(0x0A);
		$this->MD5Update ($this->valContext, $pwarray, 64);
		$this->MD5StoreDigest($this->valContext);
		$hashedsalt=str_repeat("\x00",16);
		$this->makekey(0, $key);
		$this->rc4 ($salt, 16, $key);
		$this->rc4 ($hashedsalt, 16, $key);
		$salt[16] = chr(0x80);
		$this->memset($salt, 17, chr(0), 47);
		$salt[56] = chr(0x80);
		$this->MD5Init ($mdContext2);
		$this->MD5Update ($mdContext2, $salt, 64);
		$this->MD5StoreDigest($mdContext2);
		$hashedsalt ^= $mdContext2->digest;
		return $hashedsalt;
	}

	/**
	* @access private
	*/
	function makeXorPtn($size,$pwd="", $docid="", $salt="")
	{
	if (($pwd != "") & ($docid != "") & ($salt != "")){
		$this->makehashedsalt($pwd, $docid, $salt);
	}
	if ($this->valContext->digest == str_repeat("\x00",0x10))
		die("Cannot get MD5-digest.\n");
	$key = NEW rc4_key;
	$blk = 0;
	$this->makekey($blk, $key);
	$j=0;
	$end=$size;
	$ret="";

	while ($j<$end)
		{
		$test=str_repeat("\x00",0x10);

		$this->rc4($test,0x10,$key);
		$ret .= $test;
		$j+=0x10;
		if (($j % 0x400) == 0)
			{
			$blk++;
			$this->makekey($blk, $key);
		   	}
		}
		return $ret;
	}

	/**
	* @access private
	*/
	function expandpwb ($password, &$pwarray)
	{
		$pwarray=str_repeat("\x00",64);
		$i=0;
		while(ord($password[$i])){
			$pwarray[$i*2]=$password[$i];
			$i++;
		}
		$pwarray[$i*2]=chr(0x80);
		$pwarray[56] = chr($i << 4);
	}

	/**
	* @access private
	*/
	function toDigest($src)
	{
		$digest = $this->toDigest1($src);
		return bin2hex($digest);
	}

	/**
	* @access private
	*/
	function toDigest1($src)
	{
		$md1=NEW MD5_CTX;
		$digest = str_repeat("\x00",16);

		$len = strlen($src);
		$this->MD5Init($md1);
		$this->MD5Update($md1,$src, $len);
		$this->MD5StoreDigest($md1);
//		$this->MD5Final($md1);
		return($md1->digest);
	}
}

/**
* convert UNIXTIME to MS-EXCEL time
* @param integer $timevalue UNIXTIME
* @param integer $timeoffset the hour-time difference between GMT and local time
* @return integer MS-EXCEL time
* @access public
* @example ./sample.php sample
*/
function unixtime2ms($timevalue,$timeoffset=null) {
	$offset = ($timeoffset===null)? date('Z') : $timeoffset * 3600;
	return (($timevalue + $offset) /3600 /24 + 25569);
}

/**
* convert MS-EXCEL time to UNIXTIME
* @param integer $timevalue MS-EXCEL time
* @param integer $timeoffset the hour-time difference between GMT and local time
* @return integer UNIXTIME
* @access public
*/
function ms2unixtime($timevalue,$timeoffset=null,$offset1904 = 0){
	$offset = ($timeoffset===null)? date('Z') : $timeoffset * 3600;
	if ($timevalue > 1)	$timevalue -= ($offset1904 ? 24107 : 25569);
	return round($timevalue * 24 * 3600 - $offset);
}



/**
*  Formula_Parser - A class for parsing Excel formulas
*
*  This module is adapted from Parse.php in Spreadsheet_Excel_Writer.
*  Enhanced by kishiyan <excelreviser@gmail.com> 2009.02.15
*  Maintained at http://chazuke.com/
*
*  Spreadsheet_Excel_Writer was written/ported by Xavier Noguer
*  <xnoguer@rezebra.com>.
*  He ported it from the PERL Spreadsheet::WriteExcel module.
*  The author of the Spreadsheet::WriteExcel module is John McNamara
*  <jmcnamara@cpan.org>
*
*  License Information:
*
*    This module is free software; you can redistribute it and/or
*    modify it under the terms of the GNU Lesser General Public
*    License as published by the Free Software Foundation; either
*    version 2.1 of the License, or (at your option) any later version.
*
*    This library is distributed in the hope that it will be useful,
*    but WITHOUT ANY WARRANTY; without even the implied warranty of
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
*    Lesser General Public License for more details.
*
*    You should have received a copy of the GNU General Public
*    License along with this library; if not, write to the Free Software
*    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307 USA
*
*/
define('EREG_STRING', "^\"[^\"]{0,255}\"$");
define('EREG_FUNC', "^[A-Z0-9\xc0-\xdc\.]+$");
define('PTN_FUNC', "/^[A-Z0-9\xc0-\xdc\.]+$/");
define('PTN_STRING', "/^\"[^\"]{0,255}\"$/");
define('PTN_REF_R0C0', '/^R(\d+)C(\d+)$/');
define('PTN_EXREF_R0C0', "/^S\d+(\:S\d+)?\!R\d+C\d+$/");
define('PTN_RANGE_R0C0', "/^R\d+C\d+:R\d+C\d+$/");
define('PTN_RANGE2_R0C0', "/^R\d+C\d+\.\.R\d+C\d+$/");
define('PTN_RANGE3D_R0C0', "/^S\d+(\:S\d+)?\!R\d+C\d+:R\d+C\d+$/");
define('PTN_2CELL_R0C0', "/^R\d+C\d+:R\d+C\d+$/");
define('PTN_2CELL2_R0C0', "/^R\d+C\d+\.\.R\d+C\d+$/");
// For compatibility
define('PTN_REF_A1', '/^(\$)?([A-Ia-i]?[A-Za-z])(\$)?(\d+)$/');
define('PTN_EXREF_A1', "/^\w+(\:\w+)?\![A-Ia-i]?[A-Za-z](\d+)$/");
define('PTN_EXREF2_A1', "/^'\w+(\:\w+)?'\![A-Ia-i]?[A-Za-z][0-9]+$/");
define('PTN_EX2REF_A1', "/^'\w+(\:\w+)?'\![A-Ia-i]?[A-Za-z](\d+)$/");
define('PTN_RANGE_A1', "/^(\$)?[A-Ia-i]?[A-Za-z](\$)?(\d+)\:(\$)?[A-Ia-i]?[A-Za-z](\$)?(\d+)$/");
define('PTN_RANGE2_A1', "/^(\$)?[A-Ia-i]?[A-Za-z](\$)?(\d+)\.\.(\$)?[A-Ia-i]?[A-Za-z](\$)?(\d+)$/");
define('PTN_RANGE3D_A1', "/^\w+(\:\w+)?\!([A-Ia-i]?[A-Za-z])?(\d+)\:([A-Ia-i]?[A-Za-z])?(\d+)$/");
define('PTN_RANGE3D2_A1', "/^'\w+(\:\w+)?'\!([A-Ia-i]?[A-Za-z])?(\d+)\:([A-Ia-i]?[A-Za-z])?(\d+)$/");
define('PTN_2CELL_A1', "/^([A-Ia-i]?[A-Za-z])(\d+)\:([A-Ia-i]?[A-Za-z])(\d+)$/");
define('PTN_2CELL2_A1', "/^([A-Ia-i]?[A-Za-z])(\d+)\.\.([A-Ia-i]?[A-Za-z])(\d+)$/");
define('PTN_RANG_NUMNUM', '/(\$)?(\d+)\:(\$)?(\d+)/');

class Formula_Parser
{
	// The index of the character we are currently looking at
	var $_cur_char;
	// The token we are working on.
	var $_cur_token;
	// The formula to parse
	var $_formula;
	// The character ahead of the current char
	var $_aheadchar;
	// The parse tree to be generated
	var $_parse_tree;
	// The byte order. 1 => big endian, 0 => little endian.
	var $_byte_order;
	// Array of external sheets
	var $_ext_sheets;
	// Array of sheet references in the form of REF structures
	var $_sheet_ref;
	// Array of functions in the form
	var $_functions;
	// Charset for string
	var $_charset='eucJP-win';

	/**
	* Constructor
	* $charset string for multibyte-string
	*/
	function Formula_Parser($charset='auto')
	{
		$this->_cur_char  = 0;
		$this->_cur_token = '';	   // The token we are working on.
		$this->_formula	   = "";	   // The formula to parse.
		$this->_aheadchar	 = '';	   // The character ahead of the current char.
		$this->_parse_tree	= '';	   // The parse tree to be generated.
		$this->_init();	  // Initialize the hashes: ptg's and function's ptg's
		//$this->_byte_order = $byte_order; // Little Endian or Big Endian
		$this->_ext_sheets = array();
		$this->_sheet_ref = array();
		if ((pack("N",1)==pack("L",1))) $this->_byte_order=TRUE; else $this->_byte_order= FALSE;
		if ($charset !='') $this->_charset=$charset;
	}

	/**
	* Initialize the ptg and function hashes. 
	*
	* @access private
	*/
	function _init()
	{
	// The Excel ptg indices
		$this->ptg = array(
		'ptgExp'       => 0x01,	'ptgTbl'       => 0x02,
		'ptgAdd'       => 0x03,	'ptgSub'       => 0x04,
		'ptgMul'       => 0x05,	'ptgDiv'       => 0x06,
		'ptgPower'     => 0x07,	'ptgConcat'    => 0x08,
		'ptgLT'        => 0x09,	'ptgLE'        => 0x0A,
		'ptgEQ'        => 0x0B,	'ptgGE'        => 0x0C,
		'ptgGT'        => 0x0D,	'ptgNE'        => 0x0E,
		'ptgIsect'     => 0x0F,	'ptgUnion'     => 0x10,
		'ptgRange'     => 0x11,	'ptgUplus'     => 0x12,
		'ptgUminus'    => 0x13,	'ptgPercent'   => 0x14,
		'ptgParen'     => 0x15,	'ptgMissArg'   => 0x16,
		'ptgStr'       => 0x17,	'ptgAttr'      => 0x19,
		'ptgSheet'     => 0x1A,	'ptgEndSheet'  => 0x1B,
		'ptgErr'       => 0x1C,	'ptgBool'      => 0x1D,
		'ptgInt'       => 0x1E,	'ptgNum'       => 0x1F,
		'ptgArray'     => 0x20,	'ptgFunc'      => 0x21,
		'ptgFuncVar'   => 0x22,	'ptgName'      => 0x23,
		'ptgRef'       => 0x24,	'ptgArea'      => 0x25,
		'ptgMemArea'   => 0x26,	'ptgMemErr'    => 0x27,
		'ptgMemNoMem'  => 0x28,	'ptgMemFunc'   => 0x29,
		'ptgRefErr'    => 0x2A,	'ptgAreaErr'   => 0x2B,
		'ptgRefN'      => 0x2C,	'ptgAreaN'     => 0x2D,
		'ptgMemAreaN'  => 0x2E,	'ptgMemNoMemN' => 0x2F,
		'ptgNameX'     => 0x39,	'ptgRef3d'     => 0x3A,
		'ptgArea3d'    => 0x3B,	'ptgRefErr3d'  => 0x3C,
		'ptgAreaErr3d' => 0x3D,	'ptgArrayV'    => 0x40,
		'ptgFuncV'     => 0x41,	'ptgFuncVarV'  => 0x42,
		'ptgNameV'     => 0x43,	'ptgRefV'      => 0x44,
		'ptgAreaV'     => 0x45,	'ptgMemAreaV'  => 0x46,
		'ptgMemErrV'   => 0x47,	'ptgMemNoMemV' => 0x48,
		'ptgMemFuncV'  => 0x49,	'ptgRefErrV'   => 0x4A,
		'ptgAreaErrV'  => 0x4B,	'ptgRefNV'     => 0x4C,
		'ptgAreaNV'    => 0x4D,	'ptgMemAreaNV' => 0x4E,
		'ptgMemNoMemN' => 0x4F,	'ptgFuncCEV'   => 0x58,
		'ptgNameXV'    => 0x59,	'ptgRef3dV'    => 0x5A,
		'ptgArea3dV'   => 0x5B,	'ptgRefErr3dV' => 0x5C,
		'ptgAreaErr3d' => 0x5D,	'ptgArrayA'    => 0x60,
		'ptgFuncA'     => 0x61,	'ptgFuncVarA'  => 0x62,
		'ptgNameA'     => 0x63,	'ptgRefA'      => 0x64,
		'ptgAreaA'     => 0x65,	'ptgMemAreaA'  => 0x66,
		'ptgMemErrA'   => 0x67,	'ptgMemNoMemA' => 0x68,
		'ptgMemFuncA'  => 0x69,	'ptgRefErrA'   => 0x6A,
		'ptgAreaErrA'  => 0x6B,	'ptgRefNA'     => 0x6C,
		'ptgAreaNA'    => 0x6D,	'ptgMemAreaNA' => 0x6E,
		'ptgMemNoMemN' => 0x6F,	'ptgFuncCEA'   => 0x78,
		'ptgNameXA'    => 0x79,	'ptgRef3dA'    => 0x7A,
		'ptgArea3dA'   => 0x7B,	'ptgRefErr3dA' => 0x7C,
		'ptgAreaErr3d' => 0x7D
	);

// The following hash was generated by "function_locale.pl" in the distro.
// Refer to function_locale.pl for non-English function names.
//
// The array elements are as follow:
// ptg:   The Excel function ptg code.
// args:  The number of arguments that the function takes:
//           >=0 is a fixed number of arguments.
//           -1  is a variable  number of arguments.
// class: The reference, value or array class of the function args.
// vol:   The function is volatile.
//
    $this->_functions = array(
    // func		         ptg args-min max class vol
	'COUNT'   	=> array(	0,	0,	30	),
	'IF'      	=> array(	1,	2,	3	),
	'ISNA'    	=> array(	2,	1,	1	),
	'ISERROR' 	=> array(	3,	1,	1	),
	'SUM'     	=> array(	4,	0,	30	),
	'AVERAGE' 	=> array(	5,	1,	30	),
	'MIN'     	=> array(	6,	1,	30	),
	'MAX'     	=> array(	7,	1,	30	),
	'ROW'     	=> array(	8,	0,	1	),
	'COLUMN'  	=> array(	9,	0,	1	),
	'NA'      	=> array(	10,	0,	0	),
	'NPV'     	=> array(	11,	2,	30	),
	'STDEV'   	=> array(	12,	1,	30	),
	'DOLLAR'  	=> array(	13,	1,	2	),
	'FIXED'   	=> array(	14,	2,	3	),
	'SIN'     	=> array(	15,	1,	1	),
	'COS'     	=> array(	16,	1,	1	),
	'TAN'     	=> array(	17,	1,	1	),
	'ATAN'    	=> array(	18,	1,	1	),
	'PI'      	=> array(	19,	0,	0	),
	'SQRT'    	=> array(	20,	1,	1	),
	'EXP'     	=> array(	21,	1,	1	),
	'LN'      	=> array(	22,	1,	1	),
	'LOG10'   	=> array(	23,	1,	1	),
	'ABS'     	=> array(	24,	1,	1	),
	'INT'     	=> array(	25,	1,	1	),
	'SIGN'    	=> array(	26,	1,	1	),
	'ROUND'   	=> array(	27,	2,	2	),
	'LOOKUP'  	=> array(	28,	2,	3	),
	'INDEX'   	=> array(	29,	2,	4	),
	'REPT'    	=> array(	30,	2,	2	),
	'MID'     	=> array(	31,	3,	3	),
	'LEN'     	=> array(	32,	1,	1	),
	'VALUE'   	=> array(	33,	1,	1	),
	'TRUE'    	=> array(	34,	0,	0	),
	'FALSE'   	=> array(	35,	0,	0	),
	'AND'     	=> array(	36,	1,	30	),
	'OR'      	=> array(	37,	1,	30	),
	'NOT'     	=> array(	38,	1,	1	),
	'MOD'     	=> array(	39,	2,	2	),
	'DCOUNT'  	=> array(	40,	3,	3	),
	'DSUM'    	=> array(	41,	3,	3	),
	'DAVERAGE'	=> array(	42,	3,	3	),
	'DMIN'    	=> array(	43,	3,	3	),
	'DMAX'    	=> array(	44,	3,	3	),
	'DSTDEV'  	=> array(	45,	3,	3	),
	'VAR'     	=> array(	46,	1,	30	),
	'DVAR'    	=> array(	47,	3,	3	),
	'TEXT'    	=> array(	48,	2,	2	),
	'LINEST'  	=> array(	49,	1,	4	),
	'TREND'   	=> array(	50,	1,	4	),
	'LOGEST'  	=> array(	51,	1,	4	),
	'GROWTH'  	=> array(	52,	1,	4	),
	'PV'      	=> array(	56,	3,	5	),
	'FV'      	=> array(	57,	3,	5	),
	'NPER'    	=> array(	58,	3,	5	),
	'PMT'     	=> array(	59,	3,	5	),
	'RATE'    	=> array(	60,	3,	6	),
	'MIRR'    	=> array(	61,	3,	3	),
	'IRR'     	=> array(	62,	1,	2	),
	'RAND'    	=> array(	63,	0,	0	),
	'MATCH'   	=> array(	64,	2,	3	),
	'DATE'    	=> array(	65,	3,	3	),
	'TIME'    	=> array(	66,	3,	3	),
	'DAY'     	=> array(	67,	1,	1	),
	'MONTH'   	=> array(	68,	1,	1	),
	'YEAR'    	=> array(	69,	1,	1	),
	'WEEKDAY' 	=> array(	70,	1,	2	),
	'HOUR'    	=> array(	71,	1,	1	),
	'MINUTE'  	=> array(	72,	1,	1	),
	'SECOND'  	=> array(	73,	1,	1	),
	'NOW'     	=> array(	74,	0,	0	),
	'AREAS'   	=> array(	75,	1,	1	),
	'ROWS'    	=> array(	76,	1,	1	),
	'COLUMNS' 	=> array(	77,	1,	1	),
	'OFFSET'  	=> array(	78,	3,	5	),
	'SEARCH'  	=> array(	82,	2,	3	),
	'TRANSPOSE'	=> array(	83,	1,	1	),
	'TYPE'    	=> array(	86,	1,	1	),
	'ATAN2'   	=> array(	97,	2,	2	),
	'ASIN'    	=> array(	98,	1,	1	),
	'ACOS'    	=> array(	99,	1,	1	),
	'CHOOSE'  	=> array(	100,	2,	30	),
	'HLOOKUP' 	=> array(	101,	3,	4	),
	'VLOOKUP' 	=> array(	102,	3,	4	),
	'ISREF'   	=> array(	105,	1,	1	),
	'LOG'     	=> array(	109,	1,	2	),
	'CHAR'    	=> array(	111,	1,	1	),
	'LOWER'   	=> array(	112,	1,	1	),
	'UPPER'   	=> array(	113,	1,	1	),
	'PROPER'  	=> array(	114,	1,	1	),
	'LEFT'    	=> array(	115,	1,	2	),
	'RIGHT'   	=> array(	116,	1,	2	),
	'EXACT'   	=> array(	117,	2,	2	),
	'TRIM'    	=> array(	118,	1,	1	),
	'REPLACE' 	=> array(	119,	4,	4	),
	'SUBSTITUTE'	=> array(	120,	3,	4	),
	'CODE'    	=> array(	121,	1,	1	),
	'FIND'    	=> array(	124,	2,	3	),
	'CELL'    	=> array(	125,	1,	2	),
	'ISERR'   	=> array(	126,	1,	1	),
	'ISTEXT'  	=> array(	127,	1,	1	),
	'ISNUMBER'	=> array(	128,	1,	1	),
	'ISBLANK' 	=> array(	129,	1,	1	),
	'T'       	=> array(	130,	1,	1	),
	'N'       	=> array(	131,	1,	1	),
	'DATEVALUE'	=> array(	140,	1,	1	),
	'TIMEVALUE'	=> array(	141,	1,	1	),
	'SLN'     	=> array(	142,	3,	3	),
	'SYD'     	=> array(	143,	4,	4	),
	'DDB'     	=> array(	144,	4,	5	),
	'INDIRECT'	=> array(	148,	1,	2	),
	'CLEAN'   	=> array(	162,	1,	1	),
	'MDETERM' 	=> array(	163,	1,	1	),
	'MINVERSE'	=> array(	164,	1,	1	),
	'MMULT'   	=> array(	165,	2,	2	),
	'IPMT'    	=> array(	167,	4,	6	),
	'PPMT'    	=> array(	168,	4,	6	),
	'COUNTA'  	=> array(	169,	0,	30	),
	'PRODUCT' 	=> array(	183,	0,	30	),
	'FACT'    	=> array(	184,	1,	1	),
	'DPRODUCT'	=> array(	191,	3,	3	),
	'ISNONTEXT'	=> array(	192,	1,	1	),
	'STDEVP'  	=> array(	193,	1,	30	),
	'VARP'    	=> array(	194,	1,	30	),
	'DSTDEVP' 	=> array(	195,	3,	3	),
	'DVARP'   	=> array(	196,	3,	3	),
	'TRUNC'   	=> array(	197,	1,	2	),
	'ISLOGICAL'	=> array(	198,	1,	1	),
	'DCOUNTA' 	=> array(	199,	3,	3	),
	'USDOLLAR'	=> array(	204,	1,	2	),
	'FINDB'   	=> array(	205,	2,	3	),
	'SEARCHB' 	=> array(	206,	2,	3	),
	'REPLACEB'	=> array(	207,	4,	4	),
	'LEFTB'   	=> array(	208,	1,	2	),
	'RIGHTB'  	=> array(	209,	1,	2	),
	'MIDB'    	=> array(	210,	3,	3	),
	'LENB'    	=> array(	211,	1,	1	),
	'ROUNDUP' 	=> array(	212,	2,	2	),
	'ROUNDDOWN'	=> array(	213,	2,	2	),
	'ASC'     	=> array(	214,	1,	1	),
	'JIS'     	=> array(	215,	1,	1	),
	'RANK'    	=> array(	216,	2,	3	),
	'ADDRESS' 	=> array(	219,	2,	5	),
	'DAYS360' 	=> array(	220,	2,	3	),
	'TODAY'   	=> array(	221,	0,	0	),
	'VDB'     	=> array(	222,	5,	7	),
	'MEDIAN'  	=> array(	227,	1,	30	),
	'SUMPRODUCT'	=> array(	228,	1,	30	),
	'SINH'    	=> array(	229,	1,	1	),
	'COSH'    	=> array(	230,	1,	1	),
	'TANH'    	=> array(	231,	1,	1	),
	'ASINH'   	=> array(	232,	1,	1	),
	'ACOSH'   	=> array(	233,	1,	1	),
	'ATANH'   	=> array(	234,	1,	1	),
	'DGET'    	=> array(	235,	3,	3	),
	'INFO'    	=> array(	244,	1,	1	),
	'DB'      	=> array(	247,	4,	5	),
	'FREQUENCY'	=> array(	252,	2,	2	),
	'ERROR.TYPE'	=> array(	261,	1,	1	),
	'AVEDEV'  	=> array(	269,	1,	30	),
	'BETADIST'	=> array(	270,	3,	5	),
	'GAMMALN' 	=> array(	271,	1,	1	),
	'BETAINV' 	=> array(	272,	3,	5	),
	'BINOMDIST'	=> array(	273,	4,	4	),
	'CHIDIST' 	=> array(	274,	2,	2	),
	'CHIINV'  	=> array(	275,	2,	2	),
	'COMBIN'  	=> array(	276,	2,	2	),
	'CONFIDENCE'	=> array(	277,	3,	3	),
	'CRITBINOM'	=> array(	278,	3,	3	),
	'EVEN'    	=> array(	279,	1,	1	),
	'EXPONDIST'	=> array(	280,	3,	3	),
	'FDIST'   	=> array(	281,	3,	3	),
	'FINV'    	=> array(	282,	3,	3	),
	'FISHER'  	=> array(	283,	1,	1	),
	'FISHERINV'	=> array(	284,	1,	1	),
	'FLOOR'   	=> array(	285,	2,	2	),
	'GAMMADIST'	=> array(	286,	4,	4	),
	'GAMMAINV'	=> array(	287,	3,	3	),
	'CEILING' 	=> array(	288,	2,	2	),
	'HYPGEOMDIST'	=> array(	289,	4,	4	),
	'LOGNORMDIST'	=> array(	290,	3,	3	),
	'LOGINV'  	=> array(	291,	3,	3	),
	'NEGBINOMDIST'	=> array(	292,	3,	3	),
	'NORMDIST'	=> array(	293,	4,	4	),
	'NORMSDIST'	=> array(	294,	1,	1	),
	'NORMINV' 	=> array(	295,	3,	3	),
	'NORMSINV'	=> array(	296,	1,	1	),
	'STANDARDIZE'	=> array(	297,	3,	3	),
	'ODD'     	=> array(	298,	1,	1	),
	'PERMUT'  	=> array(	299,	2,	2	),
	'POISSON' 	=> array(	300,	3,	3	),
	'TDIST'   	=> array(	301,	3,	3	),
	'WEIBULL' 	=> array(	302,	4,	4	),
	'SUMXMY2' 	=> array(	303,	2,	2	),
	'SUMX2MY2'	=> array(	304,	2,	2	),
	'SUMX2PY2'	=> array(	305,	2,	2	),
	'CHITEST' 	=> array(	306,	2,	2	),
	'CORREL'  	=> array(	307,	2,	2	),
	'COVAR'   	=> array(	308,	2,	2	),
	'FORECAST'	=> array(	309,	3,	3	),
	'FTEST'   	=> array(	310,	2,	2	),
	'INTERCEPT'	=> array(	311,	2,	2	),
	'PEARSON' 	=> array(	312,	2,	2	),
	'RSQ'     	=> array(	313,	2,	2	),
	'STEYX'   	=> array(	314,	2,	2	),
	'SLOPE'   	=> array(	315,	2,	2	),
	'TTEST'   	=> array(	316,	4,	4	),
	'PROB'    	=> array(	317,	3,	4	),
	'DEVSQ'   	=> array(	318,	1,	30	),
	'GEOMEAN' 	=> array(	319,	1,	30	),
	'HARMEAN' 	=> array(	320,	1,	30	),
	'SUMSQ'   	=> array(	321,	0,	30	),
	'KURT'    	=> array(	322,	1,	30	),
	'SKEW'    	=> array(	323,	1,	30	),
	'ZTEST'   	=> array(	324,	2,	3	),
	'LARGE'   	=> array(	325,	2,	2	),
	'SMALL'   	=> array(	326,	2,	2	),
	'QUARTILE'	=> array(	327,	2,	2	),
	'PERCENTILE'	=> array(	328,	2,	2	),
	'PERCENTRANK'	=> array(	329,	2,	3	),
	'MODE'    	=> array(	330,	1,	30	),
	'TRIMMEAN'	=> array(	331,	2,	2	),
	'TINV'    	=> array(	332,	2,	2	),
	'CONCATENATE'	=> array(	336,	0,	30	),
	'POWER'   	=> array(	337,	2,	2	),
	'RADIANS' 	=> array(	342,	1,	1	),
	'DEGREES' 	=> array(	343,	1,	1	),
	'SUBTOTAL'	=> array(	344,	2,	30	),
	'SUMIF'   	=> array(	345,	2,	3	),
	'COUNTIF' 	=> array(	346,	2,	2	),
	'COUNTBLANK'	=> array(	347,	1,	1	),
	'ISPMT'   	=> array(	350,	4,	4	),
	'DATEDIF' 	=> array(	351,	3,	3	),
	'DATESTRING'	=> array(	352,	1,	1	),
	'NUMBERSTRING'	=> array(	353,	2,	2	),
	'ROMAN'   	=> array(	354,	1,	2	),
	'GETPIVOTDATA'	=> array(	358,	2,	30	),
	'HYPERLINK'	=> array(	359,	1,	2	),
	'PHONETIC'	=> array(	360,	1,	1	),
	'AVERAGEA'	=> array(	361,	1,	30	),
	'MAXA'    	=> array(	362,	1,	30	),
	'MINA'    	=> array(	363,	1,	30	),
	'STDEVPA' 	=> array(	364,	1,	30	),
	'VARPA'   	=> array(	365,	1,	30	),
	'STDEVA'  	=> array(	366,	1,	30	),
	'VARA'    	=> array(	367,	1,	30	)
    );
}
	/**
	* Convert a token to the proper ptg value.
	*
	* @access private
	* @param mixed $token The token to convert.
	*/
	function _convert($token)
	{
		if ($token=='TRUE') return $this->_conv_bool(1);
		elseif ($token=='FALSE') return $this->_conv_bool(0);
		elseif (preg_match(PTN_STRING, $token))
		{
			return $this->_conv_string($token);
		}
		elseif (is_numeric($token))
		{
			return $this->_conv_number($token);
		}
		// match references like A1 or $A$1
		elseif (preg_match(PTN_REF_A1,$token) or
			    preg_match(PTN_REF_R0C0,$token))
		{
			return $this->_conv_ref2d($token);
		}
		// match external references like Sheet1!A1 or Sheet1:Sheet2!A1
		elseif (preg_match(PTN_EXREF_A1,$token) or
				preg_match(PTN_EXREF_R0C0,$token))
		{
			return $this->_conv_ref3d($token);
		}
		// match external references like Sheet1!A1 or Sheet1:Sheet2!A1
		elseif (preg_match(PTN_EX2REF_A1,$token))
		{
			return $this->_conv_ref3d($token);
		}
		// match ranges like A1:B2
		elseif (preg_match(PTN_RANGE_A1,$token) or
				preg_match(PTN_RANGE_R0C0,$token))
		{
			return $this->_conv_range2d($token);
		}
		// match ranges like A1..B2
		elseif (preg_match(PTN_RANGE2_A1,$token) or
				preg_match(PTN_RANGE2_R0C0,$token))
		{
			return $this->_conv_range2d($token);
		}
		// match external ranges like Sheet1!A1 or Sheet1:Sheet2!A1:B2
		elseif (preg_match(PTN_RANGE3D_A1,$token) or
				preg_match(PTN_RANGE3D_R0C0,$token))
		{
			return $this->_conv_range3d($token);
		}
		// match external ranges like 'Sheet1'!A1 or 'Sheet1:Sheet2'!A1:B2
		elseif (preg_match(PTN_RANGE3D2_A1,$token))
		{
			return $this->_conv_range3d($token);
		}
		elseif (isset($this->ptg[$token])) // operators (including parentheses)
		{
			return pack("C", $this->ptg[$token]);
		}
		// if it's an argument, ignore the token (the argument remains)
		elseif ($token == 'arg')
		{
			return '';
		}
		// TODO: use real error codes
		return $this->raiseError("Unknown token $token");
	}
	
	/**
	* Convert a boolean token to ptgBool
	*
	* @access private
	* @param integer $bool an integer for conversion to its ptg value
	*/
	function _conv_bool($bool) {
	    return pack("CC", $ptg['ptgBool'], $bool);
	}
	
	/**
	* Convert a boolean token to ptgBool
	*
	* @access private
	* @param integer $bool an integer
	*/
	function _conv_number($num)
	{
		// Integer in the range 0..2**16-1
		if ((preg_match("/^\d+$/",$num)) and ($num <= 65535)) {
			return pack("Cv", $this->ptg['ptgInt'], $num);
		}
		else // A float
		{
			if ($this->_byte_order) { // if it's Big Endian
				$num = strrev($num);
			}
			return pack("Cd", $this->ptg['ptgNum'], $num);
		}
	}
	
	/**
	* Convert a string token to ptgStr
	*
	* @access private
	* @param string $string A string for conversion to its ptg value. 
	*/
	function _conv_string($string)
	{
		// chop away beggining and ending quotes
		$string = substr($string, 1, strlen($string) - 2);
		if (strlen($string) > 255) {
			return $this->raiseError("String in formula has more than 255 chars\n");
		}
		if (mb_detect_encoding($string,"ASCII,ISO-8859-1")=="ASCII"){
			$encoding =0;
			$len = strlen($string);
		} else {
			$encoding =1;
			$string = mb_convert_encoding($string, "UTF-16LE", $this->_charset);
			$len = mb_strlen ($string,"UTF-16LE");
		}
		return pack("CCC",$this->ptg['ptgStr'], $len, $encoding).$string;
	}
 
	/**
	* Convert a function to a ptgFunc or ptgFuncVarV depending on the number of
	* args that it takes.
	*
	* @access private
	* @param string  $token	The name of the function for convertion to ptg value.
	* @param integer $num_args The number of arguments the function receives.
	* @return string The packed ptg for the function
	*/
	function _convertFunction($token, $num_args) {

	    if(isset($functions[$token][0])) return $this->raiseError("Unknown function $token() in formula\n");
			$argsmin	 = $this->_functions[$token][1];
			$argsmax	 = $this->_functions[$token][2];

	// Fixed number of args eg. TIME($i,$j,$k).
	    if ($argsmin == $argsmax) {
	    // Check that the number of args is valid.
	        if ($argsmin != $num_args) {
	            return $this->raiseError("Incorrect number of arguments for $token() in formula\n");
	        }
	        else {
	            return pack("Cv", $this->ptg['ptgFuncV'], $this->_functions[$token][0]);
	        }
	    }

	// Variable number of args eg. SUM($i,$j,$k, ..).
	    if ($argsmin > $num_args) {
	            return $this->raiseError("Less number of arguments for $token() in formula\n");
		} elseif ($argsmax < $num_args) {
	            return $this->raiseError("Too many number of arguments for $token() in formula\n");
		} else {
	        return pack("CCv", $this->ptg['ptgFuncVarV'], $num_args, $this->_functions[$token][0]);
	    }
	}
	
	/**
	* Convert an Excel range such as A1:D4 to a ptgRefV.
	*
	* @access private
	* @param string $range An Excel range in the A1:A2 or A1..A2 format.
	*/
	function _conv_range2d($range)
	{
		if (preg_match(PTN_2CELL_A1,$range)) {
			list($cell1, $cell2) = split(':', $range);
		}
		elseif (preg_match(PTN_2CELL_R0C0,$range)) {
			list($cell1, $cell2) = split(':', $range);
		}
		elseif (preg_match(PTN_2CELL2_A1,$range)) {
			list($cell1, $cell2) = split('\.\.', $range);
		}
		elseif (preg_match(PTN_2CELL2_R0C0,$range)) {
			list($cell1, $cell2) = split('\.\.', $range);
		}
		else {
			return $this->raiseError("Unknown range separator");
		}
	
		$cell_array1 = $this->_cell_to_packed_rowcol($cell1);
		if ($this->isError($cell_array1)) return $cell_array1;
		list($row1, $col1) = $cell_array1;
		$cell_array2 = $this->_cell_to_packed_rowcol($cell2);
		if ($this->isError($cell_array2)) return $cell_array2;
		list($row2, $col2) = $cell_array2;
	
		$ptgArea = pack("C", $this->ptg['ptgAreaA']);
		return $ptgArea . $row1 . $row2 . $col1. $col2;
	}
 
	/**
	* Convert an Excel 3d range such as "Sheet1!A1:D4" or "Sheet1:Sheet2!A1:D4" to
	* a ptgArea3d.
	*
	* @access private
	*/
function  _conv_range3d($token) {
		list($ext_ref, $range) = split('!', $token);
		$ext_ref = $this->_getRefIndex($ext_ref);
		if ($this->isError($ext_ref)) return $ext_ref;
		list($cell1, $cell2) = split(':', $range);
		if (!preg_match("/\d/",$cell1)) $cell1 .= '1';
		if (!preg_match("/\d/",$cell2)) $cell2 .= '65536';
		if (preg_match(PTN_REF_R0C0, $cell1))
		{
			$cell_array1 = $this->_cell_to_packed_rowcol($cell1);
			if ($this->isError($cell_array1)) return $cell_array1;
			list($row1, $col1) = $cell_array1;
			$cell_array2 = $this->_cell_to_packed_rowcol($cell2);
			if ($this->isError($cell_array2)) return $cell_array2;
			list($row2, $col2) = $cell_array2;
		} 
		elseif (preg_match(PTN_REF_A1, $cell1))
		{
			$cell_array1 = $this->_cell_to_packed_rowcol($cell1);
			if ($this->isError($cell_array1)) return $cell_array1;
			list($row1, $col1) = $cell_array1;
			$cell_array2 = $this->_cell_to_packed_rowcol($cell2);
			if ($this->isError($cell_array2)) return $cell_array2;
			list($row2, $col2) = $cell_array2;
		}
		else { // It's a rows range (like 26:27)
			 $cells_array = $this->_rangeToPackedRange($cell1.':'.$cell2);
			 if ($this->isError($cells_array)) return $cells_array;
			 list($row1, $col1, $row2, $col2) = $cells_array;
		}

		$ptgArea = pack("C", $this->ptg['ptgArea3dA']);
		return $ptgArea . $ext_ref . $row1 . $row2 . $col1. $col2;
}

	/**
	* Convert an Excel reference such as A1, $B2, C$3 or $D$4 to a ptgRefV.
	*
	* @access private
	* @param string $cell An Excel cell reference
	* @return string The cell in packed() format with the corresponding ptg
	*/
	function _conv_ref2d($cell)
	{
		$cell_array = $this->_cell_to_packed_rowcol($cell);
		if ($this->isError($cell_array)) return $cell_array;
		list($row, $col) = $cell_array;
		$ptgRef = pack("C", $this->ptg['ptgRefA']);
		return $ptgRef.$row.$col;
	}
	
	/**
	* Convert an Excel 3d reference such as "Sheet1!A1" or "Sheet1:Sheet2!A1" to a
	* ptgRef3d.
	*
	* @access private
	* @param string $cell An Excel cell reference
	*/
	function _conv_ref3d($cell)
	{
		list($ext_ref, $cell) = split('!', $cell);
		$ext_ref = $this->_getRefIndex($ext_ref);
		if ($this->isError($ext_ref)) return $ext_ref;
		list($row, $col) = $this->_cell_to_packed_rowcol($cell);
		$ptgRef = pack("C", $this->ptg['ptgRef3dA']);
		return $ptgRef . $ext_ref. $row . $col;
	}


	/**
	* Look up the REF index that corresponds to an external sheet name 
	* (or range). If it doesn't exist yet add it to the workbook's references
	* array. It assumes all sheet names given must exist.
	*
	* @access private
	* @param string $ext_ref The name of the external reference
	*/
	function _getRefIndex($ext_ref)
	{
		$ext_ref = preg_replace("/^'/", '', $ext_ref); // Remove leading  ' if any.
		$ext_ref = preg_replace("/'$/", '', $ext_ref); // Remove trailing ' if any.
 
		// Check if there is a sheet range eg., Sheet1:Sheet2.
		if (preg_match("/:/", $ext_ref))
		{
			list($sheet_name1, $sheet_name2) = split(':', $ext_ref);

			$sheet1 = $this->_getSheetIndex($sheet_name1);
			if ($sheet1 == -1) {
				return $this->raiseError("Unknown sheet name $sheet_name1 in formula");
			}
			$sheet2 = $this->_getSheetIndex($sheet_name2);
			if ($sheet2 == -1) {
				return $this->raiseError("Unknown sheet name $sheet_name2 in formula");
			}
 
			// Reverse max and min sheet numbers if necessary
			if ($sheet1 > $sheet2) {
				list($sheet1, $sheet2) = array($sheet2, $sheet1);
			}
		}
		else // Single sheet name only.
		{
			$sheet1 = $this->_getSheetIndex($ext_ref);
			if ($sheet1 == -1) {
				return $this->raiseError("Unknown sheet name $ext_ref in formula");
			}
			$sheet2 = $sheet1;
		}
		$index=$sheet1;
		return pack('v', $index);
	}

	/**
	* Look up the index that corresponds to an external sheet name. The hash of
	* sheet names is updated by the addworksheet() method of the 
	* Spreadsheet_Excel_Writer_Workbook class.
	*
	* @access private
	* @return integer The sheet index, -1 if the sheet was not found
	*/
	function _getSheetIndex($sheet_name)
	{
		if (preg_match("/^S\d+$/",$sheet_name)){
			return substr($sheet_name,1);
		}
		elseif (!isset($this->_ext_sheets[$sheet_name])) {
			return -1;
		}
		else {
			return $this->_ext_sheets[$sheet_name];
		}
	}

	/**
	* This method is used to update the array of sheet names. It is
	* called by the addWorksheet() method of the
	* Spreadsheet_Excel_Writer_Workbook class.
	*
	* @access public
	* @see Spreadsheet_Excel_Writer_Workbook::addWorksheet()
	* @param string  $name  The name of the worksheet being added
	* @param integer $index The index of the worksheet being added
	*/
	function setExtSheet($name, $index)
	{
		$this->_ext_sheets[$name] = $index;
	}

	/**
	* Convert an Excel cell reference such as A1 or $B2 or C$3 or $D$4 to a zero
	* indexed row and column number. Also returns two (0,1) values to indicate
	* whether the row or column are relative references.
	*
	* @access private
	* @param string $cell The Excel cell reference in A1 format.
	* @return array
	*/
	function _cell_to_rowcol($cell)
	{
		if (preg_match(PTN_REF_R0C0,$cell)) {
		preg_match(PTN_REF_R0C0,$cell,$match);
		// return absolute column if there is a $ in the ref
		$col_rel = 0;
		$col     = $match[2];
		$row_rel = 0;
		$row	 = $match[1];

		}
		else
		{
		preg_match(PTN_REF_A1,$cell,$match);
		// return absolute column if there is a $ in the ref
		$col_rel = empty($match[1]) ? 1 : 0;
		$col_ref = $match[2];
		$row_rel = empty($match[3]) ? 1 : 0;
		$row	 = $match[4];

		// Convert base26 column string to a number.
		$expn   = strlen($col_ref) - 1;
		$col	= 0;
		for ($i=0; $i < strlen($col_ref); $i++)
		{
			$col += (ord($col_ref{$i}) - ord('A') + 1) * pow(26, $expn);
			$expn--;
		}

		// Convert 1-index to zero-index
		$row--;
		$col--;
		}
		return array($row, $col, $row_rel, $col_rel);
	}
	

	/**
	* pack() row and column into the required 3 or 4 byte format.
	*
	* @access private
	* @param string $cell The Excel cell reference to be packed
	* @return array Array containing the row and column in packed() format
	*/
function _cell_to_packed_rowcol($cell) {

 	$cell = strtoupper($cell);
	list($row, $col, $row_rel, $col_rel) = $this->_cell_to_rowcol($cell);
	if ($col >= 256) die("Column $cell greater than IV in formula\n");
	if ($row >= 65536) die("Row $cell greater than 65536 in formula\n");

// Set the high bits to indicate if row or col are relative.
    $col    |= $col_rel << 14;
    $col    |= $row_rel << 15;

    $row     = pack('v', $row);
    $col     = pack('v', $col);

    return array($row, $col);
}
	
	/**
	* pack() row range into the required 3 or 4 byte format.
	* Just using maximum col/rows, which is probably not the correct solution
	*
	* @access private
	* @param string $range The Excel range to be packed
	* @return array Array containing (row1,col1,row2,col2) in packed() format
	*/
	function _rangeToPackedRange($range)
	{
		preg_match(PTN_RANG_NUMNUM, $range, $match);
		// return absolute rows if there is a $ in the ref
		$row1_rel = empty($match[1]) ? 1 : 0;
		$row1	 = $match[2];
		$row2_rel = empty($match[3]) ? 1 : 0;
		$row2	 = $match[4];
		// Convert 1-index to zero-index
		$row1--;
		$row2--;
		// Trick poor inocent Excel
		$col1 = 0;
		$col2 = 65535;

		if (($row1 >= 65536) or ($row2 >= 65536)) {
			return $this->raiseError("Row in: $range greater than 65536 ");
		}
	
		// Set the high bits to indicate if rows are relative.
		$col1	|= $row1_rel << 15;
		$col2	|= $row2_rel << 15;
		$col1	 = pack('v', $col1);
		$col2	 = pack('v', $col2);
		$row1	 = pack('v', $row1);
		$row2	 = pack('v', $row2);
	
		return array($row1, $col1, $row2, $col2);
	}

	/**
	* Advance to the next valid token.
	*
	* @access private
	*/
	function _advance()
	{
		$i = $this->_cur_char;
		// remove white spaces
		if ($i < strlen($this->_formula))
		{
			while ($this->_formula{$i} == " ") {
				$i++;
			}
			if ($i < strlen($this->_formula) - 1) {
				$this->_aheadchar = $this->_formula{$i+1};
			}
			$token = "";
		}
		while ($i < strlen($this->_formula))
		{
			$token .= $this->_formula{$i};
			if ($i < strlen($this->_formula) - 1) {
				$this->_aheadchar = $this->_formula{$i+1};
			}
			else {
				$this->_aheadchar = '';
			}
			if ($this->_match($token) != '')
			{
				$this->_cur_char = $i + 1;
				$this->_cur_token = $token;
				return 1;
			}
			if ($i < strlen($this->_formula) - 2) {
				$this->_aheadchar = $this->_formula{$i+2};
			}
			else {
				$this->_aheadchar = '';
			}
			$i++;
		}
		if ($this->_cur_char >= $i){
			$this->_cur_char = $i+1;
			return;
		}
//		die("Lexical error ".$this->_formula."  ".$this->_cur_char);
	}
	
	/**
	* Checks if it's a valid token.
	*
	* @access private
	* @param mixed $token The token to check.
	* @return mixed	   The checked token or false on failure
	*/
	function _match($token)
	{
		switch($token)
		{
			case '+':
			case '-':
			case '&':
			case '^':
			case '*':
			case '/':
			case '(':
			case ')':
			case ',':
			case ';':
				return $token;
				break;
			case '>':
				if ($this->_aheadchar == '=') { // it's a GE token
					break;
				}
				return $token;
				break;
			case '<':
				// it's a LE or a NE token
				if (($this->_aheadchar == '=') or ($this->_aheadchar == '>')) {
					break;
				}
				return $token;
				break;
			case '>=':
			case '<=':
			case '=':
			case '<>':
				return $token;
				break;
			default:
				// if it's a reference by R0C0
				if (preg_match(PTN_REF_R0C0,$token) and
				   !ereg("[0-9]",$this->_aheadchar) and 
				   ($this->_aheadchar != ':') and ($this->_aheadchar != '.') and
				   ($this->_aheadchar != '!') )
				{
					return $token;
				}
				// If it's an external reference (S1!R0C0 or S1:S2!R0C0)
				elseif (preg_match(PTN_EXREF_R0C0,$token) and
					   !ereg("[0-9]",$this->_aheadchar) and
					   ($this->_aheadchar != ':') and ($this->_aheadchar != '.'))
				{
					return $token;
				}
				// if it's a range (R0C0:R1C1)
				elseif (preg_match(PTN_RANGE_R0C0,$token) and 
					   !ereg("[0-9]",$this->_aheadchar))
				{
					return $token;
				}
				// if it's a range (R0C0..R0C0)
				elseif (preg_match(PTN_RANGE2_R0C0,$token) and 
					   !ereg("[0-9]",$this->_aheadchar))
				{
					return $token;
				}
				// If it's an external range like S1!R0C0:R0C0 or S1:S2!R0C0:R1C1
				elseif (preg_match(PTN_RANGE3D_R0C0,$token) and
					   !ereg("[0-9]",$this->_aheadchar))
				{
					return $token;
				}

				// if it's a reference
				elseif (preg_match(PTN_REF_A1,$token) and
				   !ereg("[0-9]",$this->_aheadchar) and 
				   ($this->_aheadchar != ':') and ($this->_aheadchar != '.') and
				   ($this->_aheadchar != '!') and ($this->_aheadchar != 'C'))
				{
					return $token;
				}
				// If it's an external reference (Sheet1!A1 or Sheet1:Sheet2!A1)
				elseif (preg_match(PTN_EXREF_A1,$token) and
					   !ereg("[0-9]",$this->_aheadchar) and
					   ($this->_aheadchar != ':') and ($this->_aheadchar != '.') 
						and ($this->_aheadchar != 'C'))
				{
					return $token;
				}
				// If it's an external reference (Sheet1!A1 or Sheet1:Sheet2!A1)
				elseif (preg_match(PTN_EXREF2_A1,$token) and
					   !ereg("[0-9]",$this->_aheadchar) and
					   ($this->_aheadchar != ':') and ($this->_aheadchar != '.'))
				{
					return $token;
				}
				// if it's a range (A1:A2)
				elseif (preg_match(PTN_RANGE_A1,$token) and 
					   !ereg("[0-9]",$this->_aheadchar))
				{
					return $token;
				}
				// if it's a range (A1..A2)
				elseif (preg_match(PTN_RANGE2_A1,$token) and 
					   !ereg("[0-9]",$this->_aheadchar))
				{
					return $token;
				}
				// If it's an external range like Sheet1!A1 or Sheet1:Sheet2!A1:B2
				elseif (preg_match(PTN_RANGE3D_A1,$token) and
					   !ereg("[0-9]",$this->_aheadchar))
				{
					return $token;
				}
				// If it's an external range like 'Sheet1'!A1 or 'Sheet1:Sheet2'!A1:B2
				elseif (preg_match(PTN_RANGE3D2_A1,$token) and
					   !ereg("[0-9]",$this->_aheadchar))
				{
					return $token;
				}
				// If it's a number (check that it's not a sheet name or range)
				elseif (is_numeric($token) and 
						(!is_numeric($token.$this->_aheadchar) or ($this->_aheadchar == '')) and
						($this->_aheadchar != '!') and ($this->_aheadchar != ':'))
				{
					return $token;
				}
				// If it's a string (of maximum 255 characters)
				elseif (ereg(EREG_STRING,$token))
				{
					return $token;
				}
				// if it's a function call
				elseif (eregi(EREG_FUNC,$token) and ($this->_aheadchar == "("))
				{
					return $token;
				}
				return '';
		}
	}
	
	/**
	* The parsing method. It parses a formula.
	*
	* @access public
	* @param string $formula The formula to parse, without the initial equal
	*						sign (=).
	*/
	function parse($formula)
	{
		$this->_cur_char = 0;
		$this->_formula = $formula;
		$this->_aheadchar = $formula{1};
		$this->_advance();
		$this->_parse_tree = $this->_cond();
		if ($this->isError($this->_parse_tree)) return $this->_parse_tree;
		return true;
	}
	
	/**
	* It parses a condition. It assumes the following rule:
	* Cond -> Expr [(">" | "<") Expr]
	*
	* @access private
	*/
	function _cond()
	{
		$result = $this->_expr();
		if ($this->isError($result)) return $result;
/*
		if ($this->_cur_token == '&')
		{
			$this->_advance();
			$result2 = $this->_expr();
			if ($this->isError($result2)) return $result2;
			$result = $this->_makeTree('ptgConcat', $result, $result2);
		} else
*/
		if ($this->_cur_token == '<')
		{
			$this->_advance();
			$result2 = $this->_expr();
			if ($this->isError($result2)) return $result2;
			$result = $this->_makeTree('ptgLT', $result, $result2);
		}
		elseif ($this->_cur_token == '>') 
		{
			$this->_advance();
			$result2 = $this->_expr();
			if ($this->isError($result2)) return $result2;
			$result = $this->_makeTree('ptgGT', $result, $result2);
		}
		elseif ($this->_cur_token == '<=') 
		{
			$this->_advance();
			$result2 = $this->_expr();
			if ($this->isError($result2)) return $result2;
			$result = $this->_makeTree('ptgLE', $result, $result2);
		}
		elseif ($this->_cur_token == '>=') 
		{
			$this->_advance();
			$result2 = $this->_expr();
			if ($this->isError($result2)) return $result2;
			$result = $this->_makeTree('ptgGE', $result, $result2);
		}
		elseif ($this->_cur_token == '=') 
		{
			$this->_advance();
			$result2 = $this->_expr();
			if ($this->isError($result2)) return $result2;
			$result = $this->_makeTree('ptgEQ', $result, $result2);
		}
		elseif ($this->_cur_token == '<>') 
		{
			$this->_advance();
			$result2 = $this->_expr();
			if ($this->isError($result2)) return $result2;
			$result = $this->_makeTree('ptgNE', $result, $result2);
		}
		return $result;
	}

	/**
	* It parses a expression. It assumes the following rule:
	* Expr -> Term [("+" | "-" | "&") Term]
	*	  -> "string"
	*	  -> "-" Term
	*
	* @access private
	*/
	function _expr()
	{
		// catch "-" Term
		if ($this->_cur_token == '-') {
			$this->_advance();
			$result2 = $this->_expr();
			if ($this->isError($result2)) return $result2;
			$result = $this->_makeTree('ptgUminus', $result2, '');
			return $result;
		}
		$result = $this->_term();
		if ($this->isError($result)) return $result;
		while (($this->_cur_token == '+') or 
			   ($this->_cur_token == '-') or
			   ($this->_cur_token == '&'))
		{
			if ($this->_cur_token == '+')
			{
				$this->_advance();
				$result2 = $this->_term();
				if ($this->isError($result2)) return $result2;
				$result = $this->_makeTree('ptgAdd', $result, $result2);
			}
			elseif ($this->_cur_token == '-')
			{
				$this->_advance();
				$result2 = $this->_term();
				if ($this->isError($result2)) return $result2;
				$result = $this->_makeTree('ptgSub', $result, $result2);
			}
			else
			{
				$this->_advance();
				$result2 = $this->_term();
				if ($this->isError($result2)) return $result2;
				$result = $this->_makeTree('ptgConcat', $result, $result2);
			}

		}
		return $result;
	}

	/**
	* It parses a term. It assumes the following rule:
	* Term -> Fact [("*" | "/") Fact]
	*
	* @access private
	*/
	function _term()
	{
		$result = $this->_term2();
		if ($this->isError($result)) return $result;
		while (($this->_cur_token == '*') or 
			   ($this->_cur_token == '/'))
		{
			if ($this->_cur_token == '*')
			{
				$this->_advance();
				$result2 = $this->_term2();
				if ($this->isError($result2)) return $result2;
				$result = $this->_makeTree('ptgMul', $result, $result2);
			}
			elseif ($this->_cur_token == '/')
			{
				$this->_advance();
				$result2 = $this->_term2();
				if ($this->isError($result2)) return $result2;
				$result = $this->_makeTree('ptgDiv', $result, $result2);
			}
		}
		return $result;
	}

	/**
	* It parses a term2. It assumes the following rule:
	* Term -> Fact [("^") Fact]
	*
	* @access private
	*/
	function _term2()
	{
		$result = $this->_fact();
		if ($this->isError($result)) return $result;
		while ($this->_cur_token == '^')
		{
				$this->_advance();
				$result2 = $this->_fact();
				if ($this->isError($result2)) return $result2;
				$result = $this->_makeTree('ptgPower', $result, $result2);
		}
		return $result;
	}

	/**
	* It parses a factor. It assumes the following rule:
	* Fact -> ( Expr )
	*	   | CellRef
	*	   | CellRange
	*	   | Number
	*	   | Function
	*
	* @access private
	*/
	function _fact()
	{
		if ($this->_cur_token == '(')
		{
			$this->_advance();		 // eat the "("
			if ($this->_cur_char >= strlen($this->_formula)) {
				return $this->raiseError("Unknown formula= $this->_formula");
			}
			$result = $this->_makeTree('ptgParen', $this->_expr(), '');
			if ($this->isError($result)) return $result;
			if ($this->_cur_token != ')') {
				return $this->raiseError("')' token expected.");
			}
			$this->_advance();		 // eat the ")"
			return $result;
		}
		// If it's a string return a string node
		if (ereg(EREG_STRING, $this->_cur_token)) {
			$result = $this->_makeTree($this->_cur_token, '', '');
			$this->_advance();
			return $result;
		} else
		// if it's a reference
		if (preg_match(PTN_REF_A1,$this->_cur_token) or
		    preg_match(PTN_REF_R0C0,$this->_cur_token))
		{
			$result = $this->_makeTree($this->_cur_token, '', '');
			$this->_advance();
			return $result;
		}
		// If it's an external reference (Sheet1!A1 or Sheet1:Sheet2!A1)
		elseif (preg_match(PTN_EXREF_A1,$this->_cur_token) or
			    preg_match(PTN_EXREF_R0C0,$this->_cur_token))
		{
			$result = $this->_makeTree($this->_cur_token, '', '');
			$this->_advance();
			return $result;
		}
		// If it's an external reference (Sheet1!A1 or Sheet1:Sheet2!A1)
		elseif (preg_match(PTN_EXREF2_A1,$this->_cur_token))
		{
			$result = $this->_makeTree($this->_cur_token, '', '');
			$this->_advance();
			return $result;
		}
		// if it's a range
		elseif (preg_match(PTN_RANGE_A1,$this->_cur_token) or 
				preg_match(PTN_RANGE2_A1,$this->_cur_token) or
				preg_match(PTN_RANGE_R0C0,$this->_cur_token) or
				preg_match(PTN_RANGE2_R0C0,$this->_cur_token))
		{
			$result = $this->_cur_token;
			$this->_advance();
			return $result;
		}
		// If it's an external range (Sheet1!A1 or Sheet1!A1:B2)
		elseif (preg_match(PTN_RANGE3D_A1,$this->_cur_token) or
			   preg_match(PTN_RANGE3D_R0C0,$this->_cur_token))
		{
			$result = $this->_cur_token;
			$this->_advance();
			return $result;
		}
		// If it's an external range ('Sheet1'!A1 or 'Sheet1'!A1:B2)
		elseif (preg_match(PTN_RANGE3D2_A1,$this->_cur_token))
		{
			$result = $this->_cur_token;
			$this->_advance();
			return $result;
		}
		elseif (is_numeric($this->_cur_token))
		{
			$result = $this->_makeTree($this->_cur_token, '', '');
			$this->_advance();
			return $result;
		}
		// if it's a function call
		elseif (eregi(EREG_FUNC,$this->_cur_token))
		{
			$result = $this->_func();
			return $result;
		}
		return $this->raiseError("Syntax error: ".$this->_cur_token.
								 ", lookahead: ".$this->_aheadchar.
								 ", current char: ".$this->_cur_char);
	}
	
	/**
	* It parses a function call. It assumes the following rule:
	* Func -> ( Expr [,Expr]* )
	*
	* @access private
	*/
	function _func()
	{
		$num_args = 0; // number of arguments received
		$function = $this->_cur_token;
		$this->_advance();
		$this->_advance();		 // eat the "("
		while ($this->_cur_token != ')')
		{
			if ($num_args > 0)
			{
				if ($this->_cur_token == ',' ||
					$this->_cur_token == ';')
				{
					$this->_advance();  // eat the "," or ";"
				}
				else {
					return $this->raiseError("Syntax error: comma expected in ".
									  "function $function, arg #{$num_args}");
				}
				$result2 = $this->_cond();
				if ($this->isError($result2)) return $result2;
				$result = $this->_makeTree('arg', $result, $result2);
			}
			else // first argument
			{
				$result2 = $this->_cond();
				if ($this->isError($result2)) return $result2;
				$result = $this->_makeTree('arg', '', $result2);
			}
			$num_args++;
		}
		if (!isset($this->_functions[$function]))
			return $this->raiseError("Unknown function \"$function()\" in formula\n");
		$argsmin	 = $this->_functions[$function][1];
		$argsmax	 = $this->_functions[$function][2];
	    if ($argsmin > $num_args) {
	            return $this->raiseError("Less number of arguments for $function() in formula\n");
		} elseif ($argsmax < $num_args) {
	            return $this->raiseError("Too many number of arguments for $function() in formula\n");
		}
	
		$result = $this->_makeTree($function, $result, $num_args);
		$this->_advance();		 // eat the ")"
		return $result;
	}
	
	/**
	* Creates a tree. In fact an array which may have one or two arrays (sub-trees)
	* as elements.
	*
	* @access private
	* @param mixed $value The value of this node.
	* @param mixed $left  The left array (sub-tree) or a final node.
	* @param mixed $right The right array (sub-tree) or a final node.
	* @return array A tree
	*/
	function _makeTree($value, $left, $right)
	{
		return array('value' => $value, 'left' => $left, 'right' => $right);
	}

	/**
	* Builds a string containing the tree in reverse polish notation (What you 
	* would use in a HP calculator stack).
	*
	* @access public
	* @param array $tree The optional tree to convert.
	* @return string The tree in reverse polish notation
	*/
	function conv_Form2bin($tree = array())
	{
		$polish = ""; // the string we are going to return
		if (empty($tree)) // If it's the first call use _parse_tree
		{
			$tree = $this->_parse_tree;
		}
		if (is_array($tree['left']))
		{
			$converted_tree = $this->conv_Form2bin($tree['left']);
			if ($this->isError($converted_tree)) return $converted_tree;
			$polish .= $converted_tree;
		}
		elseif ($tree['left'] != '') // It's a final node
		{
			$converted_tree = $this->_convert($tree['left']);
			if ($this->isError($converted_tree)) return $converted_tree;
			$polish .= $converted_tree;
		}
		if (is_array($tree['right']))
		{
			$converted_tree = $this->conv_Form2bin($tree['right']);
			if ($this->isError($converted_tree)) return $converted_tree;
			$polish .= $converted_tree;
		}
		elseif ($tree['right'] != '') // It's a final node
		{
			$converted_tree = $this->_convert($tree['right']);
			if ($this->isError($converted_tree)) return $converted_tree;
			$polish .= $converted_tree;
		}
		// if it's a function convert it here (so we can set it's arguments)
		if (preg_match(PTN_FUNC,$tree['value']) and
			!preg_match(PTN_REF_R0C0,$tree['value']) and
			!preg_match(PTN_RANGE2_R0C0,$tree['value']) and
			!preg_match(PTN_REF_A1,$tree['value']) and
			!preg_match(PTN_RANGE2_A1,$tree['value']) and
			!is_numeric($tree['value']) and
			!isset($this->ptg[$tree['value']]))
		{
			// left subtree for a function is always an array.
			if ($tree['left'] != '') {
				$left_tree = $this->conv_Form2bin($tree['left']);
			}
			else {
				$left_tree = '';
			}
			if ($this->isError($left_tree)) return $left_tree;
			// add it's left subtree and return.
			$res=$this->_convertFunction($tree['value'], $tree['right']);
			if ($this->isError($res)) return $res;
			return $left_tree.$res;
		}
		else
		{
			$converted_tree = $this->_convert($tree['value']);
			if ($this->isError($converted_tree)) return $converted_tree;
		}
		$polish .= $converted_tree;
		return $polish;
	}


	function convFormRecord($tree = array()){
//print_r($tree);
		if (empty($tree)) // If it's the first call use _parse_tree
		{
			$tree = $this->_parse_tree;
		}
		if (isset($tree['value']) and $tree['value']=='IF'){
			if ($tree['right']== 3) {
				$a0=$this->conv_Form2bin($tree['left']['left']['left']);
				if ($this->isError($a0)) return $a0;
				$a1=$this->conv_Form2bin($tree['left']['left']['right']);
				if ($this->isError($a1)) return $a1;
				$a2=$this->conv_Form2bin($tree['left']['right']);
				if ($this->isError($a2)) return $a2;
				$a3=$this->_convertFunction($tree['value'],$tree['right']);
				if ($this->isError($a3)) return $a3;
				$s3="\x19\x08".chr(strlen($a3)-1)."\x0";
				$s2="\x19\x08".chr(strlen($a2.$s3.$a3)-1)."\x0";
				$s1="\x19\x02".chr(strlen($a1.$s2))."\x0";
				$r=$a0.$s1.$a1.$s2.$a2.$s3.$a3;
			} elseif ($tree['right']== 2) {
				$a1=$this->conv_Form2bin($tree['left']['left']['right']);
				if ($this->isError($a1)) return $a1;
				$a2=$this->conv_Form2bin($tree['left']['right']);
				if ($this->isError($a2)) return $a2;
				$a3=$this->_convertFunction($tree['value'],$tree['right']);
				if ($this->isError($a3)) return $a3;
				$s3="\x19\x08".chr(strlen($a3)-1)."\x0";
				$s1="\x19\x02".chr(strlen($a2.$s3))."\x0";
				$r=$a1.$s1.$a2.$s3.$a3;
			}
			return $r;
		} elseif (isset($tree['value']) and $tree['value']=='NOW'){
			$s1="\x19\x01\x00\x00";
			$a1=$this->conv_Form2bin($tree);
			if ($this->isError($a1)) return $a1;
			return $s1.$a1;
		}

		return $this->conv_Form2bin($tree);
	}

	/**
	* @access private
	*/
	function &raiseError($a) {return new ErrMess($a);}
	/**
	* @access private
	*/
	function isError($a) {return is_a($a, 'ErrMess');}
}
?>
