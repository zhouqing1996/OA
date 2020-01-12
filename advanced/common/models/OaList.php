<?php

namespace common\models;

use Yii;

/**
 *
 * @property int $no 编号
 * @property int $listid 
 * @property varchar $listname
 * @property int $typeid
 // * @property varchar $typename

 */
class OaList extends \yii\db\ActiveRecord
{
    /**
     * {@inheritdoc}
     */
    public static function tableName()
    {
        return 'oa_list';
    }
    /**
     * {@inheritdoc}
     */
    public function rules()
    {
        return [
            [['no','listid','listname','typename','typeid'],'required'],
            [['no','listid','typeid'],'integer'],
            // [['listname','typename'],'string','max'=>255]
            [['listname'],'string','max'=>255]
        ];
    }
    /**
     * {@inheritdoc}
     */
    public function attributeLabels()
    {
        return [
            'no'=>'No',
            'listname' =>'listName',
            'listid' =>'listId',
            // 'typename' => 'typeName',
            'typeid' =>'typeId',
        ];
    }
}
