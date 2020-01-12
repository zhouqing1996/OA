<?php
namespace backend\module\document\controllers;

// require_once __DIR__.'/../../../vendor/autoload.php';
use Yii;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style;
// use PhpOffice\PhpWord\TemplateProcessor;


use yii\web\Controller;
use yii\web\Response;
use yii\web\Request;
use yii\data\Pagination;
use yii\db\Query;
use common\models\OaFlowInfo;
use common\models\OaList10;
use common\models\OaList;
use common\models\OaList1;

require_once "./word/PHPWord.php";

// <!-- 文件相关操作 -->
class FileController extends Controller
{
	public function actionIndex()
	{
		return "文件操作";
	}
	public function actionDemo()
	{
		$PHPWord = new PHPWord();
		$document = $PHPWord->loadTemplate('./word/Examples/Template.docx');
		$document->setValue('Value1',iconv('utf-8', 'GB2312//IGNORE','1'));
		$document->setValue('Value2',iconv('utf-8', 'GB2312//IGNORE','2'));
		$filename = './word/Examples/m-i-'.time().'.docx';
		$document->save($filename);
	}
	public function actionDownloaddemo()
	{
		$phpWord = new \PhpOffice\PhpWord\PhpWord();
		$section = $phpWord->addSection();
		$section->addText("Learn from yesterday, live for today, hope for tomorrow.");
    	$section->addText(
    		"Great achievement is usually born of great sacrifice.",array('name' => 'Tahoma', 'size' => 10));
    	$fontStyleName = 'oneUserDefinedStyle';
    	$phpWord->addFontStyle($fontStyleName,array('name' => 'Tahoma', 'size' => 10, 'color' => '1B2232', 'bold' => true));
    	$section->addText("The greatest accomplishment is not in never falling.",$fontStyleName);
    	$fontStyle = new \PhpOffice\PhpWord\Style\Font();
    	$fontStyle->setBold(true);
    	$fontStyle->setName('Tahoma');
    	$fontStyle->setSize(13);
    	$myTextElement = $section->addText('"Believe you can and you halfway there." (Theodor Roosevelt)');
    	$myTextElement->setFontStyle($fontStyle);
    	Header("Content-type:application/octet-stream");
    	Header("Accept-Ranges:bytes");
    	Header("Content-Disposition:attchment; filename=".'测试文件.docx')
    	ob_clean();//关键
    	flush();//关键
    	$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
    	$objWriter->save("php://output");
    	exit();
	}
	public function actionDownloadtoword()
	{
		//下载为word
		$request = \Yii::$app->request;
        $procname = $request->post('procname');
        $userid = $request->post('userid');
        $procid = $request->post('procid');
        $list = (new Query())->select('listid')->from('oa_list')->andWhere(['listname' => $procname])->one();
        if($list['listid'] == 3)
        {
            $query1 = (new Query())
                    ->select('*')
                    ->from('oa_list10')
                    ->andWhere(['userid' => $userid])
                    ->andWhere(['procid' => $procid])
                    ->andWhere(['isvaild' => 1])
                    ->one();
            return array("data" => [$query1],"msg" => "查看成功");
        }

	}
}

