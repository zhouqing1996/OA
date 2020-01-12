<?php

namespace common\models;

use Yii;

/**
 * This is the model class for table "oa_shenpi_proc".公文审批过程信息
 *
 * @property varchar $id 流水号
 * @property varchar $documentid 公文号
 * @property varchar $procid 审批流程号
 * @property varchar $stepid 步骤号
 * @property varchar $userid 用户编号
 * @property char $opinionflag 签署意见（0：不同意;1：同意;2：需进一步研究)
 * @property char $flag 签署标志 (0:未开始;1:进入审批流程;2:审批完成)
 * @property char $procflag 流程进展标志（0:审批完成;1:否定;2:发回重审）

 
 */
class OaContactInfo extends \yii\db\ActiveRecord
{
    /**
     * {@inheritdoc}
     */
    public static function tableName()
    {
        return 'oa_shenpi_proc';
    }
    /**
     * {@inheritdoc}
     */
    public function rules()
    {
        return [
            [['id', 'documentid', 'procid','stepid','userid','opinion'], 'required'],
            [['opinionflag','flag','procflag'], 'char',max => 1],
            [['id', 'documentid', 'procid','stepid','userid','opinion' ], 'varchar', 'max' => 11],
        ];
    }
    /**
     * {@inheritdoc}
     */
    public function attributeLabels()
    {
        return [
            'id' => 'ID',
            'documentid' => 'Document ID',
            'procid' => 'Proc ID',
            'stepid' => 'Step ID',
            'userid' => 'User ID',
            'opinion' => 'Opinion',
            'opinionflag' => 'Opinion Flag',
            'flag' => 'Flag',
            'procflag' => 'Proc Flag',
        ];
    }
}
