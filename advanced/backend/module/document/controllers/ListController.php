<?php
namespace backend\module\document\controllers;

use yii\web\Controller;
use yii\web\Response;
use yii\web\Request;
use yii\data\Pagination;
use yii\db\Query;
use common\models\OaList;

Class ListController extends Controller
{
	public function actionIndex()
    {
        return "分类表";
    }
    public function actionGetlistdata()
    {
        $request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->all();
        return array("data"=>$query,"msg"=>"success");
    }
    public function actionGetlist1data()
    {
        $request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>1])
            ->all();
        return array("data"=>$query,"msg"=>"success");
    }
    public function actionGetlist2data()
    {
        $request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>2])
            ->all();
        return array("data"=>$query,"msg"=>"success");
    }
    public function actionGetlist3data()
    {
        $request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>3])
            ->all();
        return array("data"=>$query,"msg"=>"success");
    }
    public function actionGetlistnewdata()
    {
        $request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->all();
        return array("data"=>$query,"msg"=>"success");
    }
    // public function actionGetlistdata()
    // {

    //     $request = \Yii::$app->request;
    //     $query = (new Query())
    //         ->select('*')
    //         ->from('oa_list')
    //         ->all();
            
    //     $query1 = (new Query())
    //         ->select('*')
    //         ->from('oa_list');

    //     $countQuery = clone $query1;
    //     $totalCount = $countQuery->count();
    //     $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
    //     $models = $query1->offset($pages->offset)
    //         ->limit($pages->limit)
    //         ->all();   
    //     $pageNum = ceil($totalCount/8);
    //     return array("data"=>[$query,$pageNum],"msg"=>"success");
    // }
    public function actionGetlist1()
    {
    	$request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>1])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>1]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionGetlist2()
    {
    	$request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>2])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>2]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionGetlist3()
    {
    	$request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>3])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>3]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionGetlist4()
    {
    	$request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>4])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>4]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionGetlist5()
    {
    	$request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>5])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>5]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionGetlist6()
    {
    	$request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>6])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>6]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionGetlist7()
    {
    	$request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>7])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>7]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionGetlist8()
    {
    	$request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>8])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_list')
            ->andWhere(['typeid' =>8]);

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