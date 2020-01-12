<?php

namespace backend\module\home\controllers;

use yii\web\Controller;
use yii\helpers\ArrayHelper;
use common\models\OaContactInfo;
/**
 * Default controller for the `home` module
 */
class IndexController extends Controller
{
    /**
     * Renders the index view for the module
     * @return string
     */
    public function actionIndex()
    {
        return "首页";
    }
    public function actionLogin()
    {
    	$request = \Yii::$app->request;
        $id = $request->post('username');
        $password = $request->post('password');

        $result = OaContactInfo::find()
                    ->andWhere(['id' =>$id])
                    ->andWhere(['password' =>$password])
                    ->andWhere(['status' =>1])
                    ->one();
        if($result===null){
        	return array("data"=>"11","msg"=>"登陆失败");
        }
        else{
        	return array("data"=>$result,"msg"=>"登陆成功");
        }
    }
    public function actionLogout()
    {
        
    }
    
}
