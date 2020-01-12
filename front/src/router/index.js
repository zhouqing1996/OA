import Home from '../pages/home/Index.vue'
import Content from '../pages/home/Content.vue'
import Login from '../pages/home/Login.vue'

import Doc from '../pages/document/Index'
import Mycontact from '../pages/document/Mycontact'
import Myfile from '../pages/document/Myfile'
import Flownew from '../pages/document/Flownew'
import Flowschedule from '../pages/document/Flowschedule'
import Flowmy from '../pages/document/Flowmy'
import Flowok from "../pages/document/Flowok";

import List1 from "../pages/document/List/List1";
import List2 from "../pages/document/List/List2";
import List10 from "../pages/document/List/List10";
import Listnew from "../pages/document/List/Listnew";

import Looklist10 from "../pages/document/LookList/Looklist10";
import Looklist2 from "../pages/document/LookList/Looklist2";
import Looklist1 from "../pages/document/LookList/Looklist1";

import Test from "../pages/test/Test";

export default [
  {
    path: '/',
    redirect: '/home'
  },
  {
    path: '/document/index',
    name:'docu',
    component:Doc
  },
  {
    path:'/document/list/list1',
    name:'list1',
    component:List1
  },
  {
    path:'/document/list/list2',
    name:'list2',
    component:List2
  },
  {
    path:'/document/list/list10',
    name:'list10',
    component:List10
  },
  {
    path:'/document/list/listnew',
    name:'listnew',
    component:Listnew
  },
  {
    path:'/document/LookList/Looklist1',
    name:'looklist1',
    component:Looklist1
  },
  {
    path:'/document/LookList/Looklist2',
    name:'looklist2',
    component:Looklist2
  },
  {
    path:'/document/LookList/Looklist10',
    name:'looklist10',
    component:Looklist10
  },
  {
    path:'//test/Test',
    name:'test',
    component:Test
  },
  {
    path: '/home',
    redirect: '/home/login',
    component: Home,
    children: [
      {
        path: 'index',
        component: Content
      },
      {
        path:'content',
        component:Content
      },
      {
        path: 'login',
        component: Login
      }
    ]
  },
  {
    path: '/document/index',
    name: 'Doc',
    component: Doc,
    children: [
      {
        // 通讯录
        path: '/document/mycontact',
        name: 'mycontact',
        component: Mycontact
      },
      {
        // 个人文件柜
        path: '/document/myfile',
        name:'myfile',
        component:Myfile
      },
      {
        // 新建流程
        path: '/document/flownew',
        name:'flownew',
        component:Flownew
      },
      {
        // 待办流程
        path: '/document/flowschedule',
        name:'flowschedule',
        component:Flowschedule
      },
      {
        // 我的流程
        path:'/document/flowmy',
        name:'flowmy',
        component:Flowmy
      },
      {
        // 已审流程
        path:'/document/flowok',
        name:'flowok',
        component:Flowok
      },

    ]
  }
]
