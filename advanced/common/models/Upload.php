<?php

namespace common\models;
use yii\base\Model;
use yii\web\UploadedFile;

class Upload extends Model
{
	public $file;

    public function rules()
    {
        return [
            [['file'], 'file', 'skipOnEmpty' => false, 'extensions' => 'png,jpg,docx'],
        ];
    }
    
    public function upload()
    {
        if ($this->validate()) {
        	$new_name = 'uploads/' . $this->file->baseName . '-' .time(). '.' . $this->file->extension;
        	$this->file->saveAs($new_name);
        	return $new_name;   
        } else {
        	return false;
        }
    }
	// public $File;
	// public function rules(){
	// 	return [
	// 		[['File'],'file','skipOnEmpty'=>false,'extensions'=>'pdf,docx,doc,jpg,png'],
	// 	];
	// }
	// public function upload()
	// {
	// 	if($this->validate())
	// 	{
	// 		$this->File->saveAs('uploads/'.$this->File->baseName.'.'.$this->File->extension);
	// 		return true;
	// 	}
	// 	else
	// 	{
	// 		retun false;
	// 	}
	// }
}