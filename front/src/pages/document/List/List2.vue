<template>
  <!--  网络与新媒体审批备案申请-->
  <div class="div1">
    <br>
    <span class="span1">优秀共青团员（干部）申请</span>
    <br/>
    <div>
      <table :model="Listform">
        <tr>
          <td class="td1">申请人</td>
<!--          <td>{{Listform.user}}</td>-->
          <td>{{getUserName}}</td>
          <td class="td1">申请部门</td>
<!--          <td>{{Listform.userDer}}</td>-->
          <td>{{getUserDpt}}</td>
          <td class="td1">申请时间</td>
          <td>{{Listform.createtime}}</td>
        </tr>
        <tr>
          <td class="td1" colspan="1">审批内容</td>
          <td colspan="6">
            <input type="text" v-model="Listform.c">
          </td>
        </tr>
        <tr>
          <td class="td1">审批人(一级审批)</td>
          <td colspan="2">
            <select v-model="Listform.p1" style="font-size: 15px;width: 180px">
              <option :value="item.id" v-for="item in contactList1" >{{item.name}}({{item.zhiwei}})</option>
            </select>
          </td>
          <td class="td1">审批人（二级审批)</td>
          <td colspan="2">
            <select v-model="Listform.p2" style="font-size: 15px;width: 180px">
              <option :value="item.id" v-for="item in contactList2">{{item.name}}({{item.zhiwei}})</option>
            </select>
          </td>
        </tr>
      </table>
    </div>
    <div class="div2">
      <button class="btn1" type="submit" v-on:click="submit(Listform)">提交申请</button>
      <button class="btn1" v-on:click="stop">终止申请</button>
<!--      <button class="btn1" v-on:click="download">下载</button>-->
    </div>
    <br/>
    <span class="span2"> 你(<span style="color: red">{{getUserName}}</span>)正在申请优秀共青团员（干部）申请</span>
  </div>
</template>

<script>
    import docxtemplater from 'docxtemplater';
    import PizZip from 'pizzip';
    import JSZipUtils from 'jszip-utils';
    import JSZip from 'jszip';

    import {saveAs} from 'file-saver';
    export default {
        name: "List2",
        data(){
            return{
                // user1:this.$store.getters.getUsername,
                Listform:{
                    procname:"优秀共青团员（干部）申请",
                    user:this.$store.getters.getUsername,
                    userid:this.$store.getters.getUserid,
                    userDer:this.$store.getters.getUserdpt,
                    createtime:new Date(),
                    p1:'',//一级审批人
                    p2:'',//二级审批人
                    c:"",//审核内容
                },
                contactList1:[],
                contactList2:[],
                file:'',
            }
        },
        methods:{
            download:function(){
                let that = this;
                JSZipUtils.getBinaryContent("../../../../static/List2.docx", function(error, content) {
                    if(error)
                    {
                        throw error;
                    }
                    let zip = new PizZip(content);
                    let doc = new docxtemplater().loadZip(zip);
                    doc.setData({
                        ...that.Listform.procname,
                        table:that.Listform
                    });
                    try {
                        doc.render();
                    }catch (error ){
                        let e = {
                            message:error.message,
                            name:error.name,
                            stack:error.stack,
                            properties: error.properties
                        };
                        console.log(JSON.stringify({error:e}));
                        throw error;
                    }
                    let out = doc.getZip().generate({
                        type:"blob",
                        mimeType:
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    });
                    saveAs(out,"List2.docx");
                });
            },
            submit:function(Listform) {
                // 提交申请
                console.log(Listform);
                this.Date();
                let that = this;
                if (that.Listform.procname != '' && that.Listform.userid!= ''  )
                    this.$http.post('/document/flow/newflow', {
                        procname: that.Listform.procname,
                        username: that.Listform.user,
                        userid: that.Listform.userid,
                        userDpt: that.Listform.name,
                        createtime:that.Listform.createtime,
                        stepid1: that.Listform.p1,
                        stepid2: that.Listform.p2,
                        content1: that.Listform.c,
                        text:that.Listform.t,
                    }).then(res =>{
                        console.log(res);
                        alert("提交申请成功！");
                        this.$router.go(-1);
                    }).catch(function (err) {
                        console.log(err);
                    })
                else{
                    alert('提交申请失败！');
                }
            },
            stop:function(){
                // 终止申请
                this.$router.go(-1);
            },
            Date:function(){
                var d = new Date();
                var year = d.getFullYear();
                var dateArr =[d.getMonth()+1,d.getDate(),d.getHours(),d.getMinutes(),d.getSeconds()];
                for(var i=0;i<dateArr.length;i++)
                {
                    if(dateArr[i]>=1&&dateArr[i]<=9){
                        dateArr[i]="0"+dateArr[i];
                    }
                }
                var strD = year+'-'+dateArr[0]+'-'+dateArr[1]+' '+dateArr[2]+':'+dateArr[3]+':'+dateArr[4];
                this.Listform.createtime = strD;
                console.log(this.Listform.createtime);
            },
            getContactList1:function(){
                // 联系人列表
                let that = this;
                this.$http.post('/document/contact/getdata1',{
                    userdpt:this.$store.getters.getUserdpt,
                    user:this.$store.getters.getUserid,
                }).then(function (res) {
                    console.log(res.data.data)
                    that.contactList1 = res.data.data[0];
                })
            },
            getContactList2:function(){
                // 联系人列表
                let that = this;
                this.$http.post('/document/contact/getdata2',{
                    // userdpt:this.$store.getters.getUserdpt,
                    user:this.$store.getters.getUserid,
                }).then(function (res) {
                    console.log(res.data.data)
                    that.contactList2 = res.data.data[0];
                })
            },

        },
        created() {
            this.Date();
            this.getContactList1();
            this.getContactList2();
        },
        computed:{
            getUserName(){
                return this.$store.getters.getUsername
            },
            getUserId() {
                return this.$store.getters.getUserid
            },
            getUserDpt() {
                return this.$store.getters.getUserdpt
            }
        }
    }
</script>

<style scoped>
  @import "../../../common/css/list.css";

</style>
