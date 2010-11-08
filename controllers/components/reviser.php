<?php

App::import('Vendor', 'reviser/reviser');

// reviserディレクトリのパス
define("REVISER_PATH", APP . 'vendors/reviser/');
// reviserテンプレート配置用のパス
define("REVISER_TEMPLATE_PATH", REVISER_PATH . 'template/');

class ReviserComponent extends Object {
    // controller
    private $controller;
    // reviser
    private $reviser;
    
    function initialize(&$controller) {
        $this->controller =& $controller;
    }

    function startup(&$controller) {
    }
    
    /**
     * リバイザ初期化
     * 
     * @access public
     * @author sakuragawa
     */
    public function create(){
        $this->reviser = new Excel_Reviser();
        
        // エンコーディング指定
        $enc = Configure::read('App.encoding');
        $this->reviser->setInternalCharset($enc);
        
        $this->_limitOff();
    }
    
    
    /**
     * リバイザ出力
     * 
     * @access public
     * @author sakuragawa
     * 
     * @param $reviser リバイザ オブジェクト
     * @param $template リバイザテンプレート
     * @param $outFile 出力ファイル名(nullの場合はリバイザテンプレート名を使用する[default:null])
     * @param $path サーバローカルに出力する場合のパス(nullの場合はWeb出力[default:null])
     */
    public function outPut($template, $outFile = null, $path = null){
        if(is_null($outFile)){
            $count = substr_count($template, '/');
            if($count != 0){
                // テンプレートがパスになっている
                
                // ファイル名のみを使用する
                $arr = explode('/', $template);
                $outFile = $arr[$count];
            }else{
                // テンプレートファイル名である
                
                // 出力ファイル名にテンプレートファイル名を使用する
                $outFile = $template;
            }
        }
        
        // テンプレートパス
        $readfile = sprintf("%s%s", REVISER_TEMPLATE_PATH, $template);

        // 出力
        $this->reviser->reviseFile($readfile, $outFile, $path);
    }
    
    
    /**
     * addStringのラッパー
     * 
     * @access public
     * @author sakuragawa
     */
    public function addString($sheet, $row, $col, $str, $refrow=null, $refcol=null, $refsheet=null){
        $this->reviser->addString($sheet, $row, $col, $str, $refrow, $refcol, $refsheet);
    }
    
    /**
     * addNumberのラッパー
     * 
     * @access public
     * @author sakuragawa
     */
    public function addNumber($sheet, $row, $col, $str, $refrow=null, $refcol=null, $refsheet=null){
        $this->reviser->addNumber($sheet, $row, $col, $str, $refrow, $refcol, $refsheet);
    }
    
    
    /**
     * copyRowLineのラッパー
     * 
     * @access public
     * @author sakuragawa
     */
    public function copyRowLine($sheetsrc, $rowsrc, $sheetdest, $rowdest, $num, $inc=0){
        $this->reviser->copyRowLine($sheetsrc, $rowsrc, $sheetdest, $rowdest, $num, $inc);
    }
    
    
    /**
     * addFormulaのラッパー
     * 
     * @access public
     * @author sakuragawa
     */
    public function addFormula($sheet, $row, $col, $formula, $refrow=null, $refcol=null, $refsheet=null, $opt=0){
        $this->reviser->addFormula($sheet, $row, $col, $formula, $refrow, $refcol, $refsheet, $opt);
    }
    
    
    /**
     * setSheetnameのラッパー
     * 
     * @access public
     * @author sakuragawa
     */
    public function setSheetname($sn,$str){
        $this->reviser->setSheetname($sn,$str);
    }
    
    
    /**
     * addSheetのラッパー
     * 
     * @access public
     * @author sakuragawa
     */
    public function addSheet($orgsn, $num){
        $this->reviser->addSheet($orgsn, $num);
    }
    
    
    /**
     * rmSheetのラッパー
     * 
     * @access public
     * @author sakuragawa
     */
    public function rmSheet($sn){
        $this->reviser->rmSheet($sn);
    }
    
    /**
     * 帳票出力等でメモリリミット、タイムアウト等の制限を解除する
     * 
     * @access private
     * @author sakuragawa
     */
    private function _limitOff(){
        // メモリ制限解除
        ini_set('memory_limit',-1);
        // タイムアウト制限を無しに設定
        set_time_limit(0);
        
        if(Configure::read('debug') != 0){
            return ;
        }
        
        // エラー出力なし(ワーニングさえでない。)
        error_reporting(0);
    }
    /*
     $reviser = new Excel_Reviser();
     $reviser->setInternalCharset('utf-8');
     */
}

?>