<?php

namespace common\models;

use Yii;

/**
 * This is the model class for table "major_info".
 *
 * @property int $id 联系人的id
 * @property varchar $name 姓名
 * @property int $sex 性别
 * @property varchar $department 部门（学院）
 * @property varchar $position 职位
 * @property int $phone 
 * @property varchar $zhiwei 行政职位
 * @property varchar $email
 * @property int $status
 * @property int $kind 等级限制 1 管理员 2学生

 * @property User[] $users


 */
class OaContactInfo extends \yii\db\ActiveRecord
{
    /**
     * {@inheritdoc}
     */
    public static function tableName()
    {
        return 'oa_contact';
    }
    /**
     * {@inheritdoc}
     */
    public function rules()
    {
        return [
            [['id', 'name','password','department','kind','zhiwei'], 'required'],
            [['status'], 'integer'],
            [['department', 'position', ], 'string', 'max' => 50],
            [['sex'], 'string', 'max' => 2],
            [['name','password','zhiwei'], 'string', 'max' => 255],
            [['phone','kind','id'],'string','max'=>11]
        ];
    }
    /**
     * {@inheritdoc}
     */
    public function attributeLabels()
    {
        return [
            'id' => 'ID',
            'name' => 'Name',
            'password' => 'Password',
            'kind' => 'Kind',
            'sex' => 'Sex',
            'department' => 'Department',
            'position' => 'Position',
            'phone' => 'Phone',
            'email' => 'Email',
            'zhiwei' =>'Zhiwei',
            'status' => 'Status',
        ];
    }
    /**
     * @return \yii\db\ActiveQuery
     */
    public function findByID()
    {
        return static::findOne(['id' => $id, 'status' => self::STATUS_ACTIVE]);
    }
}
