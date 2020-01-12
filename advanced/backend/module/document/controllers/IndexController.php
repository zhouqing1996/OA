<?php

namespace backend\module\document\controllers;

use yii\web\Controller;

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
        return "办公系统首页";
    }
}
