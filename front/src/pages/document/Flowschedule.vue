<template>
<!--    待办流程-->
  <div class="display1">
    <div class="display2">
      <table>
        <tr>
          <th>序号</th>
          <th>流程名称</th>
          <th>发起人</th>
          <th>发起时间</th>
          <th>操作</th>
        </tr>
        <tr v-for="(n,key) in List">
          <td>{{key+1}}</td>
          <td>{{n.procname}}</td>
          <td v-if="n.proc_userid==item.id" v-for="item in contactList">{{item.name}}
          </td>
          <td>{{n.proc_createtime}}</td>
          <td>
            <button v-if="n.procname == '网络与新媒体审批备案申请'">
              <router-link :to="{name:'looklist10',query:{userid:n.proc_userid,procid:n.procid}}" >查看</router-link>
            </button>
            <button v-if="n.procname == '师范生教师专业能力训练、竞赛申请（推荐）表'">
              <router-link :to="{name:'looklist1',query:{userid:n.proc_userid,procid:n.procid}}" >查看</router-link>
            </button>
            <button v-if="n.procname == '优秀共青团员（干部）申请'">
              <router-link :to="{name:'looklist2',query:{userid:n.proc_userid,procid:n.procid}}" >查看</router-link>
            </button>
            <button  v-on:click="Wok(n.procid)">审批</button>
            <button  v-on:click="Wrefuse(n.procid)">拒绝</button>
          </td>
        </tr>
      </table>
    </div>
  </div>
</template>

<script>
    export default {
        name: "Schedule",
        data(){
            return{
                List:[],
                userid:this.$store.getters.getUserid,
                contactList:[],

            }
        },
        methods:{
            getFlow:function () {
                // 流程中审批人
                console.log(this.userid);
                this.$http.post('/document/flow/shenpiflow', {
                    user:this.userid
                }).then(res => {
                    console.log(res.data);
                    this.List = res.data.data[0];
                }).catch(function (error) {
                    console.log(error)
                })
            },
            getData:function(){
                let that = this;
                this.$http.get('/document/contact/getdata?page=1').then(function (res) {
                    console.log(res.data.data)
                    that.contactList = res.data.data[0];
                }).catch(function (err) {
                    console.log(err);
                })
            },
            Wrefuse:function (procid) {
                //拒绝同意
                this.$http.post('/document/flow/shenpirefuse',{
                    // userid:userid,
                    procid:procid,//查找到对应的流程
                    id:this.userid
                }).then(res =>{
                    console.log(res);
                    this.getFlow();
                }).catch(function (err) {
                    console.log(err)
                })
            },
            Wok:function (id) {
                //审批同意
                this.$http.post('/document/flow/shenpiok',{
                    userid:this.userid,
                    procid:id,
                }).then(res =>{
                    console.log(res);
                    this.getFlow();
                }).catch(function (err) {
                    console.log(err)
                })
            },
        },
        created() {
            this.getData();
            this.getFlow();
        },
    }
</script>

<style scoped>
@import "../../common/css/oa.css";
</style>
