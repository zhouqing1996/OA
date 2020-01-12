<?php

namespace common\models;

use Yii;

/**
 *
 * @property int $userid 用户编号
 * @property varchar $procname 流程名 
 * @property int $procid 流程id
 * @property data $createtime 流程创建时间
 * @property int $stepid1 
 * @property varchar $step1
 * @property int $stepid2 
 * @property varchar $step2
 * @property int $stepid3 
 * @property varchar $step3
 * @property char $procflag 完成标志
 * @property char $isvaild
 * @property varchar $content 内容
 * @property varchar $text 上传文件

 */
class OaList10 extends \yii\db\ActiveRecord
{
    /**
     * {@inheritdoc}
     */
    public static function tableName()
    {
        return 'oa_list10';
    }
    /**
     * {@inheritdoc}
     */
    public function rules()
    {
        return [
            [['userid','procname','procid','createtime','procflag','isvaild'],'required'],
            [['userid','procid','stepid1','stepid2','stepid3'],'integer','max' => 11],
            [['procname','step1','step2','step3','content','text'],'string','max'=>255],
            [['createtime'],'datetime'],
            [['procflag','isvaild'],'char','max' =>1],
        ];
    }
    /**
     * {@inheritdoc}
     */
    public function attributeLabels()
    {
        return [
            'userid'=>'Userid',
            'procname' =>'Procname',
            'procid' =>'Procid',
            'createtime' => 'Createtime',
            'text' => 'Text',
            'content' => 'Content',
            'procflag' =>'Procflag',
            'stepid1' => 'Stepid1',
            'step1' => 'Step1',
            'stepid2' => 'Stepid2',
            'step2' => 'Step2',
            'stepid3' => 'Stepid3',
            'step3' => 'Step3',
            'isvaild' => 'Isvaild'
        ];
    }
}