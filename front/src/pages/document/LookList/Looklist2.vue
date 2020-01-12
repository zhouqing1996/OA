<template>
  <div class="div1">
    <span class="span1">{{procname}}</span>
    <br>
    <br/>
    <div>
      <table >
        <tr>
          <td class="td1">申请人</td>
          <td><span v-if="Listform.userid==item.id" v-for="item in contactList">{{item.name}}</span></td>
          <td class="td1">申请部门</td>
          <td><span v-if="Listform.userid==item.id" v-for="item in contactList">{{item.department}}</span></td>
          <td class="td1">申请时间</td>
          <td>{{Listform.createtime}}</td>
        </tr>
        <tr>
          <td class="td1" colspan="2">审批内容</td>
          <td colspan="4">
            <span>{{Listform.content}}</span>
          </td>
        </tr>
        <tr>
          <td class="td1">审批人(一级审批)</td>
          <td colspan="2">
            <span v-if="Listform.stepid1==item.id" v-for="item in contactList">{{item.name}}({{item.zhiwei}})</span>
          </td>
          <td class="td1">审批意见</td>
          <td colspan="2">
            <span v-if="Listform.step1==0">未审批</span>
            <span v-if="Listform.step1==1">已审批</span>
            <span v-if="Listform.step1==2">已拒绝</span>
          </td>
        </tr>
        <tr>
          <td class="td1">审批人（二级审批)</td>
          <td colspan="2">
            <span v-if="Listform.stepid2==item.id" v-for="item in contactList">{{item.name}}({{item.zhiwei}})</span>
          </td>
          <td class="td1">审批意见</td>
          <td colspan="2">
            <span v-if="Listform.step2==0">未审批</span>
            <span v-if="Listform.step2==1">已审批</span>
            <span v-if="Listform.step2==2">已拒绝</span>
          </td>
        </tr>
      </table>
    </div>
    <br/>
    <div class="div1">
      <button class="btn2" type="submit" v-on:click="download">下载文件</button>
      <button class="btn2" v-on:click="back">返回</button>
      <br/>
      <span class="span2"> 你(<span style="color: red">{{user1}}</span>)正在优秀共青团员（干部）申请</span>
    </div>
  </div>
</template>

<script>
    import docxtemplater from 'docxtemplater';
    import PizZip from 'pizzip';
    import JSZipUtils from 'jszip-utils';
    import {saveAs} from 'file-saver';
    export default {
        name: "Looklist2",
        data(){
            return{
                Listform:[],
                contactList:[],
                user1:this.$store.getters.getUsername,
                procname:'优秀共青团员（干部）申请'
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
                        ...that.procname,
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
            getlooklistdata:function () {
                //获得列表10数据
                this.$http.post('/document/flow/looklistdata',{
                    procname:this.procname,
                    userid:this.$route.query.userid,
                    procid:this.$route.query.procid
                }).then(res =>{
                    this.Listform = res.data.data[0];
                    console.log(res.data.data);
                }).catch(function (err) {
                    console.log(err)
                })
            },
            getContactData:function(){
                let that = this;
                this.$http.get('/document/contact/getdata?page=1').then(function (res) {
                    console.log(res.data.data)
                    that.contactList = res.data.data[0];
                }).catch(function (err) {
                    console.log(err);
                })
            },
            back:function () {
                this.$router.go(-1);
            }
        },
        created() {
            this.getlooklistdata();
            this.getContactData();
        },
        watch:{
            '$route'(to,from){
                this.getlooklistdata();
            }
        }
    }
</script>

<style scoped>
  @import "../../../common/css/list.css";
</style>
