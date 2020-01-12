<?php
namespace backend\module\document\controllers;

use Yii;
use yii\web\Controller;
use yii\web\Response;
use yii\web\Request;
use yii\data\Pagination;
use yii\db\Query;
use common\models\OaFlowInfo;
use common\models\OaList10;
use common\models\OaList;
use common\models\OaList1;
use common\models\Upload;
use yii\web\UploadedFile;

class FlowController extends Controller
{
    public function actionIndex()
    {
        return "文件流转";
    }
    public function actionGetdata()
    {

        $request = \Yii::$app->request;
        $query = (new Query())
            ->select('*')
            ->from('oa_flow')
            ->all();

        $query1 = (new Query())
            ->select('*')
            ->from('oa_flow');
        return array("data"=>[$query],"msg"=>"success");
    }
    public function actionNewflow()
    {
        // 新建流程

        $id = (new Query())
            ->select('*')
            ->from('oa_flow')
            ->max('procid');

        $request = \Yii::$app->request;
        $procid = $id + 1;
        $procname = $request->post('procname');
        $username = $request->post('username');
        $userid = $request->post('userid');
        $userdpt = $request->post('userDpt');
        $createtime = $request->post('createtime');
        $stepid1 = $request->post('stepid1');
        $stepid2 = $request->post('stepid2');
        $content = $request->post('content1');
        if($userid == $stepid1 || $userid == $stepid2 || $stepid1 == $stepid2)
        {
            return "审批人不能是申请本人";
        }
        $result1 = \Yii::$app->db->createCommand()->insert('oa_flow',array(
            'procid' => $procid,
            'procname' =>$procname,
            'proc_userid' => $userid,
            'proc_stepid1' => $stepid1,
            'step1' => 0,
            'proc_stepid2' => $stepid2,
            'step2' => 0,
            'proc_ok' => 0,
            'isvaild' => 1,
            'proc_createtime' => $createtime)
        )->execute();

        $list = (new Query())->select('listid')->from('oa_list')->andWhere(['listname' => $procname])->one();

        if($list['listid'] == 2 && $result1 == 1)
        {
            $result2 = \Yii::$app->db->createCommand()->insert('oa_list10',array(
                'userid' => $userid,
                'procname' => $procname,
                'procid' => $procid,
                'createtime' => $createtime,
                'content' => $content,
                'procflag' => 0,
                'stepid1' => $stepid1, 
                'step1' => 0,
                'stepid2' => $stepid2,
                'step2' => 0,
                'isvaild' =>1)
            )->execute();
            if($result2==1)
            {
                return array("data" => [$result2],"msg" =>"流程创建成功");   
            }
            else
            {
                return "流程创建失败";
            }
        }
        else
        {
            return "失败";
        }


        if($list['listid'] == 10 && $result1 == 1)
        {
            $result2 = \Yii::$app->db->createCommand()->insert('oa_list10',array(
                'userid' => $userid,
                'procname' => $procname,
                'procid' => $procid,
                'createtime' => $createtime,
                'content' => $content,
                'procflag' => 0,
                'stepid1' => $stepid1, 
                'step1' => 0,
                'stepid2' => $stepid2,
                'step2' => 0,
                'isvaild' =>1)
            )->execute();
            if($result2==1)
            {
                return array("data" => [$result2],"msg" =>"流程创建成功");   
            }
            else
            {
                return "流程创建失败";
            }
        }
        else
        {
            return "失败";
        }
    }
    public function actionTest1()
    {
        $request = \Yii::$app->request;
        $content = $request->post('doc');
        return array("data" => $content,"msg" =>"test");
        // $model = new Upload();
        // if(Yii::$app->request->isPost)
        // {
        //     $model->File = UploadedFile::getInstance($model,$content);
        //     if($model->upload())
        //     {
        //         // 文件上传成功
        //         return;
        //     }
        // }
        // return array("data" => $model,"msg" =>"test");

    }
    public function actionTest2()
    {
        $model = new Upload();
        if(Yii::$app->request->isPost)
        {
            // $mfile = \Yii::$app->request->post('doc');
            $model->file = UploadedFile::getInstanceByName('doc');
            $fileName = $model->upload();
            if($fileName)
            {
                return $fileName;
            }
            else
            {
                return '文件无法保存';
            }
        }
        else
        {
            return '非法访问';
        }
        //     $model->File = UploadedFile::getInstance($model,$mfile);
        //     if($model->upload())
        //     {
        //         return;
        //     }
        // }
        // return $this->render('upload',['model' => $model]);
        // return array("data" => $model,"msg" =>"test");
    }
    public function actionNewflow1()
    {
        // 新建流程,文件转

        $id = (new Query())
            ->select('*')
            ->from('oa_flow')
            ->max('procid');

        $request = \Yii::$app->request;
        $procid = $id + 1;
        $procname = $request->post('procname');
        $username = $request->post('username');
        $userid = $request->post('userid');
        $userdpt = $request->post('userDpt');
        $createtime = $request->post('createtime');
        $stepid1 = $request->post('stepid1');
        $stepid2 = $request->post('stepid2');
        $content = $request->post('text');//文件路径


        if($userid == $stepid1 || $userid == $stepid2 || $stepid1 == $stepid2)
        {
            return "审批人不能是申请本人";
        }
        $result1 = \Yii::$app->db->createCommand()->insert('oa_flow',array(
            'procid' => $procid,
            'procname' =>$procname,
            'proc_userid' => $userid,
            'proc_stepid1' => $stepid1,
            'step1' => 0,
            'proc_stepid2' => $stepid2,
            'step2' => 0,
            'proc_ok' => 0,
            'isvaild' => 1,
            'proc_createtime' => $createtime)
        )->execute();

        $list = (new Query())->select('listid')->from('oa_list')->andWhere(['listname' => $procname])->one();

        if($list['listid'] == 1 && $result1 == 1)
        {
            $result2 = \Yii::$app->db->createCommand()->insert('oa_list10',array(
                'userid' => $userid,
                'procname' => $procname,
                'procid' => $procid,
                'createtime' => $createtime,
                'content' => $content,
                'procflag' => 0,
                'stepid1' => $stepid1, 
                'step1' => 0,
                'stepid2' => $stepid2,
                'step2' => 0,
                'isvaild' =>1)
            )->execute();
            if($result2==1)
            {
                return array("data" => [$result2],"msg" =>"流程创建成功");   
            }
            else
            {
                return "流程创建失败";
            }
        }
        else
        {
            return "失败";
        }
    }
    public function actionDownload()
    {
        // 下载申请表

    }
    public function actionMyflow()
    {
        //我的流程，包括我创建的
        $request = \Yii::$app->request;
        // 我创建的
        $userid = $request->post('user');
        $query = (new Query())
                ->select('*')
                ->from('oa_flow')
                ->andWhere(['proc_userid' => $userid])
                ->andWhere(['isvaild' => 1])
                ->all();
        return array("data" =>[$query],"msg" => "查找成功");
    }
    public function actionMycancel()
    {
        // 撤销
        $request = \Yii::$app->request;
        $procid = $request->post('procid');
        $procname = $request->post('procname');

        $result1 = \Yii::$app->db->createCommand()->update('oa_flow',
            ['isvaild' => 0],'procid=:procid',[':procid' => $procid])->execute();

        $list = (new Query())->select('listid')->from('oa_list')->andWhere(['listname' => $procname])->one();

        if($list['listid']==10 && $result1 == 1)
        {
            $result2 = \Yii::$app->db->createCommand()->update('oa_list10',
                ['isvaild' => 0],'procid=:procid',[':procid' => $procid])->execute();
            if($result2==1)
            {
                return array("data" => [$result1,$result2],"msg" => "撤销成功");
            }
            else
            {
                return "撤销失败";
            }
        }
        
    }
    public function actionShenpiflow()
    {
        $request = \Yii::$app->request;
        // 需要我审批的,满足要求，第一审批人或者是经过审批的第二审批人
        $stepid = $request->post('user');
        $query1 = (new Query())
                ->select('*')
                ->from('oa_flow')
                ->andWhere(['proc_stepid1' => $stepid])
                ->andWhere(['step1' => 0])
                ->andWhere(['proc_ok'=> 0])
                ->andWhere(['isvaild' => 1])
                ->all();
        $query2 = (new Query())
                ->select('*')
                ->from('oa_flow')
                ->andWhere(['proc_stepid2' => $stepid])
                ->andWhere(['step1' => 1])
                ->andWhere(['proc_ok'=> 0])
                ->andWhere(['step2' => 0])
                ->andWhere(['isvaild' => 1])
                ->all();
        $query = $query1+$query2;

        return array("data" =>[$query],"msg" => "查找成功");
    }
    public function actionLooklistdata()
    {
        //获得申请表的内容
        $request = \Yii::$app->request;
        $procname = $request->post('procname');
        $userid = $request->post('userid');
        $procid = $request->post('procid');
        $list = (new Query())->select('listid')->from('oa_list')->andWhere(['listname' => $procname])->one();
        if($list['listid'] == 2)
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
        if($list['listid'] == 10)
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
    public function actionShenpirefuse()
    {
        //拒绝审批
        $request = \Yii::$app->request;
        $procid = $request->post('procid');
        $id = $request->post('userid');
        $query1 = (new Query())
                ->select('proc_stepid1')
                ->from('oa_flow')
                ->andWhere(['procid'=> $procid])
                ->andWhere(['proc_stepid1' => $id])
                ->andWhere(['step1' => 0])
                ->andWhere(['isvaild' => 1])
                ->andWhere(['proc_ok' => 0])
                ->one();
        if($query1 ===null)
        {
            return array("data" => [],"msg" =>"拒绝");
        }
        else
        {
            $result1 = \Yii::$app->db->createCommand()->update('oa_flow',
                ['step1' => 2,'proc_ok'=>2],'procid=:procid',[':procid' => $procid])->execute();
            $result2 = \Yii::$app->db->createCommand()->update('oa_list10',
                ['step1' => 2,'procflag'=>2],'procid=:procid',[':procid' => $procid])->execute(); 
            if($result1==1 && $result2 == 1)
            {
                return array("data" => [$query1,$result1,$result2],"msg" =>"拒绝成功1");
            }
            else
            {
                return array("data" => [$query1,$result1,$result2],"msg" => "拒绝失败1");
            }
        }
        $query2 = (new Query())
                ->select('proc_stepid2')
                ->from('oa_flow')
                ->andWhere(['procid'=> $procid])
                ->andWhere(['proc_stepid2' => $id])
                ->andWhere(['step1' => 1])
                ->andWhere(['isvaild' => 1])
                ->andWhere(['step2' => 0])
                ->andWhere(['proc_ok' => 0])
                ->one();
        if($query2 === null){
            return array("data" => [],"msg" =>"拒绝");
        }
            else
        {
            $result1 = \Yii::$app->db->createCommand()->update('oa_flow',
                ['step2' => 2,'proc_ok'=>2],'procid=:procid',[':procid' => $procid])->execute();
            $result2 = \Yii::$app->db->createCommand()->update('oa_list10',
                ['step2' => 2,'procflag'=>2],'procid=:procid',[':procid' => $procid])->execute(); 
            if($result1==1 && $result2 == 1)
            {
                return array("data" => [$query1,$result1,$result2],"msg" =>"拒绝成功2");
            }
            else
            {
                return array("data" => [$query1,$result1,$result2],"msg" => "拒绝失败2");
            }
        }
    }
    public function actionShenpiok()
    {
        // 同意审批
        $request = \Yii::$app->request;
        $procid = $request->post('procid');
        $id = $request->post('userid');
        $query1 = (new Query())
                ->select('proc_stepid1')
                ->from('oa_flow')
                ->andWhere(['procid'=> $procid])
                ->andWhere(['proc_stepid1' => $id])
                ->andWhere(['step1' => 0])
                ->andWhere(['isvaild' => 1])
                ->andWhere(['proc_ok' => 0])
                ->one();
        if($query1['proc_stepid1']==$id)
        {
            $result1 = \Yii::$app->db->createCommand()->update('oa_flow',
                ['step1' => 1],'procid=:procid',[':procid' => $procid])->execute();
            $result2 = \Yii::$app->db->createCommand()->update('oa_list10',
                ['step1' => 1],'procid=:procid',[':procid' => $procid])->execute(); 
            if($result1==1 && $result2 == 1)
            {
                return array("data" => [$query1,$result1,$result2],"msg" =>"审批成功1");
            }
            else
            {
                return array("data" => [$query1,$result1,$result2],"msg" => "审批失败1");
            }
        }
        $query2 = (new Query())
                ->select('proc_stepid2')
                ->from('oa_flow')
                ->andWhere(['procid'=> $procid])
                // ->andWhere(['proc_stepid2' => $id])
                ->andWhere(['step1' => 1])
                ->andWhere(['isvaild' => 1])
                ->andWhere(['step2' => 0])
                ->andWhere(['proc_ok' => 0])
                ->one();
        if($query2['proc_stepid2']==$id)
        {
            $result1 = \Yii::$app->db->createCommand()->update('oa_flow',
                ['step2' => 1,'proc_ok'=>1],'procid=:procid',[':procid' => $procid])->execute();
            $result2 = \Yii::$app->db->createCommand()->update('oa_list10',
                ['step2' => 1,'procflag'=>1],'procid=:procid',[':procid' => $procid])->execute(); 
            if($result1==1 && $result2 == 1)
            {
                return array("data" => [$query1,$result1,$result2],"msg" =>"审批成功2");
            }
            else
            {
                return array("data" => [$query1,$result1,$result2],"msg" => "审批失败2");
            }
        }
        
    }
}