<?php

namespace common\models;

use Yii;

/**
 * This is the model class for table "oa_flow".流程信息
 *
 * @property varchar $procid  流程编号
 * @property varchar $procname  流程名称
 * @property varchar $proc_userid  创建流程人id
 * @property varchar $proc_stepid1  审批流程人id
 * @property varchar $step1 审核意见 0未审批，1已审批，2已拒绝
 * @property varchar $proc_stepid2  审批流程人id
 * @property varcher $step2
 * @property varchar $proc_ok  流程完成标志 0 未完成，1 成功，2 被拒
 * @property varchar $proc_createtime  流程时间
 * @property varchar $isvaild   生效标志，0：未生效，1：已生效.

 */
class oaflowinfo extends \yii\db\ActiveRecord
{
    /**
     * {@inheritdoc}
     */
    public static function tableName()
    {
        return 'oa_flow';
    }

    /**
     * {@inheritdoc}
     */
    public function rules()
    {
        return [
            [['procid', 'procname', 'proc_userid','proc_ok','isvaild','proc_createtime'], 'required'],
            [['procid','proc_userid','proc_stepid1','proc_stepid2'], 'integer', 'max' => 11],
            [['procname','step1','step2'], 'varchar', 'max' => 255],
            [['isvaild','proc_ok'], 'char', 'max' => 1],
            [['proc_createtime'],'datatime'],
        ];
    }

    /**
     * {@inheritdoc}
     */
    public function attributeLabels()
    {
        return [
            'procid' => 'ProcID',
            'procnamer' => 'ProcName',
            'proc_userid' =>'ProcUserid',
            'proc_stepid1' =>'ProcStepid1',
            'step1' => 'Step1',
            'proc_stepid2' =>'ProcStepid2',
            'step2' => 'Step2',
            'proc_ok' => 'ProcOk',
            'isvaild' => 'IsVaild',
            'proc_createtime' =>'ProcCreatetime'
        ];
    }
