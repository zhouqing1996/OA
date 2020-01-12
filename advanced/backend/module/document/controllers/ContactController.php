<?php
namespace backend\module\document\controllers;

use yii\web\Controller;
use yii\web\Response;
use yii\web\Request;
use yii\data\Pagination;
use yii\db\Query;
use common\models\OaContactInfo;

class ContactController extends Controller
{
    public function actionIndex()
    {
        return "通讯录";
    }
    public function actionGetdata()
    {

        $request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_contact')
            ->andWhere(['status' => 1])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_contact')
            ->andWhere(['status' => 1]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionGetdata1()
    {
        // 第一个审批人

        $request = \Yii::$app->request;
        $userdpt = $request->post('userdpt');
        $user = $request->post('user');
        $query = (new Query())
            ->select('*')
            ->from('oa_contact')
            // ->where('and','department' => $userdpt,['not in','id',$user])
            ->andWhere(['department'=> $userdpt])
            ->andWhere(['!=','id',$user])
            ->andWhere(['status' => 1])
            ->all();
            
        $query1 = (new Query())
            ->select('*')
            ->from('oa_contact')
            ->andWhere(['status' => 1]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionGetdata2()
    {
        // 第二个审批人

        $request = \Yii::$app->request;
        // $userdpt = $request->post('userdpt');
        $user = $request->post('user');
        $query = (new Query())
            ->select('*')
            ->from('oa_contact')
            ->andWhere(['zhiwei' =>'校长'])
            ->orWhere(['zhiwei' =>'副校长'])
            ->orWhere(['zhiwei' =>'书记'])
            ->andWhere(['!=','id',$user])
            ->andWhere(['status' => 1])
            ->all();
            
        // $query1 = (new Query())
        //     ->select('*')
        //     ->from('oa_contact')
        //     ->andWhere(['status' => 1]);

        // $countQuery = clone $query1;
        // $totalCount = $countQuery->count();
        // $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        // $models = $query1->offset($pages->offset)
        //     ->limit($pages->limit)
        //     ->all();   
        // $pageNum = ceil($totalCount/8);
        return array("data"=>[$query],"msg"=>"success");
    }
    public function actionSearch()
    {
        $request = \Yii::$app->request;
        $contactName = $request->post('contactName');
        $query = (new Query())
            ->select('*')
            ->from('oa_contact')
            ->andWhere(['name' => $contactName])
            ->andWhere(['status' => 1])
            ->all();
        $query1 = (new Query())
            ->select('*')
            ->from('oa_contact')
            ->andWhere(['status' => 1]);

        $countQuery = clone $query1;
        $totalCount = $countQuery->count();
        $pages = new Pagination(['totalCount' => $countQuery->count(),'pageSize'=>'8']);
        $models = $query1->offset($pages->offset)
            ->limit($pages->limit)
            ->all();   
        $pageNum = ceil($totalCount/8);
        return array("data"=>[$query,$pageNum],"msg"=>"success");
    }
    public function actionInsert()
    {
        $request = \Yii::$app->request;
        $Name = $request->post('name');
        $Id = $request->post('id');
        $Sex = $request->post('sex');
        $Kind = $request->post('kind');
        $Phone = $request->post('phone');
        $Department =$request->post('department');
        $Position = $request->post('position');
        $Email = $request->post('email');

        $result1 = OaContactInfo::find()
                    ->andWhere(['id' =>$Id])
                    ->andWhere(['status' =>1])
                    ->one();
        if($result1==null){
            $reslut2 =\Yii::$app->db->createCommand()->insert('oa_contact',
                array(
                    'id' => $Id,
                    'name'=>$Name,
                     'sex'=>$Sex,
                     'kind' =>$Kind,
                     'phone'=>$Phone,
                     'department' =>$Department,
                     'position' =>$Position,
                     'email' =>$Email,
                     'password' =>"123456",
                     'status'=>1)
            )->execute();
            return array("data"=>[$reslut2],"msg"=>"添加成功");
        }
        else{
            return "已存在此id";
        }
    }
    public function actionDelete()
    {
        $request = \Yii::$app->request;
        $Id = $request->post('id');
        $result = \Yii::$app->db->createCommand()
            ->update('oa_contact',['status'=>0],'id=:Id',
            [':Id' => $Id])->execute();
        return $result;
    }
    public function actionChange()
    {
        $request = \Yii::$app->request;
        $Id = $request->post('id');
        $result = OaContactInfo::find()
                    ->andWhere(['id' =>$Id])
                    ->one();
        // ->where(['id' => $Id])->one();
        if ($result==null) {
            return "failure";
        } else {
            return array("data"=>$result,"msg"=>"success");
        }
    }
    public function actionUpdate()
    {
        $request = \Yii::$app->request;
        $Name = $request->post('name');
        $Id = $request->post('id');
        $Sex = $request->post('sex');
        $Kind = $request->post('kind');
        $Phone = $request->post('phone');
        $Department =$request->post('department');
        $Position = $request->post('position');
        $Email = $request->post('email');

        $result = \Yii::$app->db->createCommand()->update('oa_contact',
            [
                'id' => $Id,
                'name'=>$Name,
                'sex'=>$Sex,
                'kind' => $Kind,
                'phone'=>$Phone,
                'department' =>$Department,
                'position' =>$Position,
                'email' =>$Email],'id=:Id',[':Id' => $Id])->execute();
        return array("data"=>$result,"msg"=>"success");
    }
}