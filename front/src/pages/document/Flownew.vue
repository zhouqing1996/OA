<template>
<!--    新建流程-->
  <div class="display1">
    <button class="btn3" v-on:click="class1" v-bind:class="{active:isActive}">全部</button>
    <button class="btn3" v-on:click="class2" v-bind:class="{active:isActive1}">创优争先</button>
    <button class="btn3" v-on:click="class3" v-bind:class="{active:isActive3}">共青团</button>
    <button class="btn3" v-on:click="class4" v-bind:class="{active:isActive3}">网络信息</button>
    <button class="btn3" v-on:click="classnew" v-bind:class="{active:isActivenew}">
      <router-link :to="{name:'listnew'}">新建流程表</router-link>
    </button>

    <div class="display2" v-show="c">
      <table >
        <tr>
          <th>序号 </th>
          <th>流程名称 </th>
          <th>流程类型</th>
          <th>操作 </th>
        </tr>
        <tr v-for=" (member,key) in List" >
          <td>{{key+1}}</td>
          <td>{{member.listname}}</td>
          <td>{{member.typename}}</td>
          <td>
            <button class="btn3" v-if="member.listid==1">
              <router-link :to="{name:'list1'}"><i class="el-icon-edit" ></i>新建</router-link>
            </button>
            <button class="btn3" v-if="member.listid==2">
              <router-link :to="{name:'list2'}" ><i class="el-icon-edit" ></i>新建</router-link>
            </button>
            <button class="btn3" v-if="member.listid==3">
              <router-link :to="{name:'list10'}" target="_blank"><i class="el-icon-edit" ></i>新建</router-link>
            </button>
          </td>
        </tr>
      </table>
    </div>

    <div class="display2" v-show="c1">
      <table >
        <tr>
          <th>序号 </th>
          <th>流程名称 </th>
          <th>流程类型</th>
          <th>操作 </th>
        </tr>
        <tr v-for=" (member,key) in List1" >
          <td>{{key+1}}</td>
          <td>{{member.listname}}</td>
          <td>{{member.typename}}</td>
          <td>
            <button class="btn3" v-if="member.listid==1">
              <router-link :to="{name:'list1'}"><i class="el-icon-edit" ></i>新建</router-link>
            </button>
          </td>
        </tr>
      </table>
    </div>

    <div class="display2" v-show="c2">
      <table>
        <tr>
          <th>序号 </th>
          <th>流程名称 </th>
          <th>操作 </th>
        </tr>
        <tr v-for=" (member,key) in List2" >
          <td>{{key+1}}</td>
          <td>{{member.listname}}</td>
          <td>
            <button class="btn3" v-if="member.listid==2">
              <router-link :to="{name:'list2'}" ><i class="el-icon-edit" ></i>新建</router-link>
            </button>
          </td>
        </tr>
      </table>
    </div>

    <div class="display2" v-show="c3">
      <table>
        <tr>
          <th>序号 </th>
          <th>流程名称 </th>
          <th>操作 </th>
        </tr>
        <tr v-for=" (member,key) in List3" >
          <td>{{key+1}}</td>
          <td>{{member.listname}}</td>
          <td>
            <button class="btn3" v-if="member.listid==3">
              <router-link :to="{name:'list10'}">新建</router-link>
            </button>
          </td>
        </tr>
      </table>
    </div>

    <div class="display2" v-show="cnew">

    </div>
  </div>
</template>

<script>
    export default {
        name: "Newfile",
        data(){
            return{
                isActive:true,
                c:true,//1全部
                isActive1:false,
                c1:false,//1全部
                isActive2:false,
                c2:false,//2
                isActive3:false,
                c3:false,//3
                isActivenew:false,
                cnew:false,//新建流程表

                List:[],
                List1:[],
                List2:[],
                List3:[],
            }
        },
        methods:{
            class1:function () {
                this.isActive=true;
                this.c=true;
                this.isActive1=false;
                this.c1=false;
                this.isActive2=false;
                this.c2=false;
                this.isActive3=false;
                this.c3=false;
                this.isActivenew=false;
                this.cnew=false;
            },
            class2:function () {
                this.isActive=false;
                this.c=false;
                this.isActive1=true;
                this.c1=true;
                this.isActive2=false;
                this.c2=false;
                this.isActive3=false;
                this.c3=false;
                this.isActivenew=false;
                this.cnew=false;
            },
            class3:function () {
                this.isActive=false;
                this.c=false;
                this.isActive1=false;
                this.c1=false;
                this.isActive2=true;
                this.c2=true;
                this.isActive3=false;
                this.c3=false;
                this.isActivenew=false;
                this.cnew=false;
            },
            class4:function () {
                this.isActive=false;
                this.c=false;
                this.isActive1=false;
                this.c1=false;
                this.isActive2=false;
                this.c2=false;
                this.isActive3=true;
                this.c3=true;
                this.isActivenew=false;
                this.cnew=false;
            },
            classnew:function () {
                this.isActive=false;
                this.c=false;
                this.isActive1=false;
                this.c1=false;
                this.isActive2=false;
                this.c2=false;
                this.isActive3=false;
                this.c3=false;
                this.isActivenew=true;
                this.cnew=true;
            },
            getListData:function() {
                //全部表
                let that = this;
                this.$http.get('/document/list/getlistdata?page=1').then(function (res) {
                    console.log(res.data.data);
                    that.List = res.data.data;
                })
            },
            getList1:function () {
                let that =this;
                this.$http.get('/document/list/getlist1data?page=1').then(function (res) {
                    console.log(res.data.data);
                    that.List1 = res.data.data;
                })
            },
            getList2:function () {
                let that =this;
                this.$http.get('/document/list/getlist2data?page=1').then(function (res) {
                    console.log(res.data.data);
                    that.List2 = res.data.data;
                })
            },
            getList3:function () {
                let that =this;
                this.$http.get('/document/list/getlist3data?page=1').then(function (res) {
                    console.log(res.data.data);
                    that.List3= res.data.data;
                })
            },

        },
        created() {
            this.getListData();
            this.getList1();
            this.getList2();
            this.getList3();
        }
    }
</script>

<style scoped>
@import "../../common/css/oa.css";
</style>
