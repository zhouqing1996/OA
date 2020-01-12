<template>
  <div style=" margin-top:20px;">
<!--    <div class="meauPanel">-->
<!--      <span>-->
<!--        <a target="_blank" href="http://phpitem.grouptong.top:8081/ScormPlatform/learnPlat/portal/index.html" >教学资源共享平台<i class="el-icon-goods" style="font-size: 55px"></i></a>-->
<!--      </span>-->
<!--    </div >-->
<!--    <div class="meauPanel">-->
<!--      <span>-->
<!--        <a target="_blank" href="http://phpitem.grouptong.top:8081/trainManage_lvliang/trainPlatform/department/major/major?name=%E9%99%A2%E7%B3%BB%E7%AE%A1%E7%90%86" >校企合作平台<i class="iconbiyeshixi" style="font-size: 55px"></i></a>-->
<!--      </span>-->
<!--    </div >-->
<!--    <div class="meauPanel">-->
<!--      <span>-->
<!--        <a target="_blank" href="http://phpitem.grouptong.top:8081/smartQ/smartQ-master/question/mywj.html">智能问卷调查系统<i class="icongongneng_fabuxuqiu" style="font-size: 55px"></i></a>-->
<!--      </span>-->
<!--    </div>-->
<!--    <div class="meauPanel">-->
<!--      <span>-->
<!--        <a target="_blank" href="http://phpitem.grouptong.top:8081/graduate_sys_thinkphp/" >研究生管理平台<i class="el-icon-edit-outline-->
<!--" style="font-size: 55px"></i></a>-->
<!--      </span>-->
<!--    </div >-->
<!--    <div class="meauPanel">-->
<!--      <span>-->
<!--        <a target="_blank" href="http://phpitem.grouptong.top:8081/resourceMaking/splitScreen-master/splitSelect.html">资源编辑系统<i class="icontiaochawenjuan1" style="font-size: 50px"></i></a>-->
<!--      </span>-->
<!--    </div>-->
    <div class="meauPanel">
      <span>
        <router-link :to="{name:'docu'}">办公OA<i class="el-icon-message" style="font-size: 50px"></i></router-link>
      </span>
    </div>

    <div class="meauPanel">
      <span>
        <a href="#" @click="toUrl('排课系统')">排课系统<i class="el-icon-date" style="font-size: 50px"></i></a>
      </span>
    </div>
    <div class="meauPanel">
      <span>
        <router-link :to="{name:'test'}">尝试<i class="el-icon-message" style="font-size: 50px"></i></router-link>
      </span>
    </div>
    <div class="clean"></div>
  </div>
</template>

<script>
    import Header from "../../components/Header";
    import docxtemplater from 'docxtemplater';
    import PizZip from 'pizzip';
    import JSZipUtils from 'jszip-utils';
    import {saveAs} from 'file-saver';
    export default {
        components: {Header},
        data () {
            return {
            }
        },
        methods: {
            alertMsg () {
                alert('无权限查看')
            },
            toUrl (title) {
                this.$alert('正在建设中……', title, {
                    confirmButtonText: '确定',
                    callback: action => {
                    }
                })
            }
        },
     }
</script>

<style scoped>
  .meauPanel{
    text-align: center;
    line-height: 260px;
    width: 250px;
    height: 230px;
    float: left;
    border: 1px solid #ECECEC;
    background:#FFF;
    /*border:1px solid #fff;*/
    border-radius: 10px;
    font-size: 18px;
    margin: 10px;
  }
</style>
