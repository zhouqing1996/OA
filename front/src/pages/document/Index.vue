<template>
  <div >
    <el-col :span="5">
      <el-menu :unique-opened="true"
               :default-active="$route.path"
               @select="handleSelect"
               style="background-color: white">
<!--        <el-submenu index="docu1">-->
<!--          <template slot="title"><i class="el-icon-user"></i><span>我的</span></template>-->
<!--          <el-menu-item-group >-->
            <el-menu-item index="mycontact"><i class="el-icon-user"></i>通讯录</el-menu-item>
<!--            <el-menu-item index="myfile">个人文件柜</el-menu-item>-->
<!--          </el-menu-item-group>-->
<!--        </el-submenu>-->

<!--        <el-submenu index="docu2">-->
<!--          <template slot="title"><i class="el-icon-folder"></i><span>流程审批</span></template>-->
<!--          <el-menu-item-group>-->
            <el-menu-item index="flownew"><i class="el-icon-folder-add"></i>新建流程</el-menu-item>
            <el-menu-item index="flowmy"><i class="el-icon-s-custom"></i>我的流程</el-menu-item>
            <el-menu-item index="flowschedule"><i class="el-icon-message-solid"></i>待办流程</el-menu-item>

<!--            <el-menu-item index="flowok">已审流程</el-menu-item>-->
<!--          </el-menu-item-group>-->
<!--        </el-submenu>-->

      </el-menu>
    </el-col>
    <el-col :span="17">
      <router-view></router-view>
    </el-col>
  </div>
</template>

<script>
    export default {
        name: "Index",
        data() {
            return {

            };
        },
        methods: {
            handleSelect(path) {
                this.$router.push(path)
            },
        }
    }
</script>
<style scoped>
@import "../../common/css/oa.css";
</style>
