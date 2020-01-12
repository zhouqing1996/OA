<?php
namespace backend\module\document\controllers;

use yii\web\Controller;
use yii\web\Response;
use yii\web\Request;
use yii\data\Pagination;
use yii\db\Query;
use common\models\OaFlowprocInfo;

class ContactController extends Controller
{
    public function actionIndex()
    {
        return "流程步骤";
    }
    public function actionGetdata()
    {

        $request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_flowproc')
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_flowproc');

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
}