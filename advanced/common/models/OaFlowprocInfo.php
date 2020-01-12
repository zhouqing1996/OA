<?php

namespace common\models;

use Yii;

/**
 * This is the model class for table "oa_flowproc".流程步骤库
 *
 * @property varchar $procid 流程编号
 * @property varchar $stepid 数据编号
 * @property varchar $userid 签署人id
 * @property char $flag 签署标志：0--签署：1—会签


 */
class OaContactInfo extends \yii\db\ActiveRecord
{
    /**
     * {@inheritdoc}
     */
    public static function tableName()
    {
        return 'oa_flowproc';
    }
    /**
     * {@inheritdoc}
     */
    public function rules()
    {
        return [
            [['procid', 'stepid', 'userid','flag'], 'required'],
            [['procid', 'stepid','userid' ], 'varchar', 'max' => 11],
            [['flag'], 'char', 'max' => 1],
        ];
    }
    /**
     * {@inheritdoc}
     */
    public function attributeLabels()
    {
        return [
            'procid' => 'Proc ID',
            'stepid' => 'Step ID',
            'userid' => 'User ID',
            'flag' => 'flag',
]        ];
    }
}
