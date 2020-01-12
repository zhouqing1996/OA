<template>
<!--    通讯录-->
  <div class="display1">
    <div class="display2" v-show="self1">
      <div class="s1">
        <div class="s2">
          <el-input v-model="inputName" placeholder="请输入通讯录名字" size="mini"></el-input>
        </div>
        <button class="btn2 el-icon-search" v-on:click="searchContact">搜索</button>
        <button class="btn2 el-icon-circle-plus-outline" v-if="kind == 1" v-on:click="dialogflag1=true" >添加</button>
        <el-dialog title="添加":visible.sync="dialogflag1" >
          <el-form :model="Listform">
            <el-form-item label="编号" :label-width="formLabelWidth">
              <el-input style="width: auto" v-model="Listform.id" auto-complete="off"></el-input>
            </el-form-item>
            <el-form-item label="姓名" :label-width="formLabelWidth">
              <el-input style="width: auto" v-model="Listform.name" auto-complete="off"></el-input>
            </el-form-item>
            <el-form-item label="性别" :label-width="formLabelWidth">
              <select v-model="Listform.sex" style="font-size: 15px;width: auto;height:auto ">
                <option value="1">男</option>
                <option value="2">女</option>
              </select>
            </el-form-item>
            <el-form-item label="身份" :label-width="formLabelWidth">
              <select v-model="Listform.kind" style="font-size: 15px;width: auto;height:auto ">
                <option value="1">管理员</option>
                <option value="2">学生</option>
              </select>
            </el-form-item>
            <el-form-item label="职位" :label-width="formLabelWidth">
              <el-input style="width: auto" v-model="Listform.position" auto-complete="off"></el-input>
            </el-form-item>
            <el-form-item label="部门(学院)" :label-width="formLabelWidth">
              <el-input style="width: auto" v-model="Listform.department" auto-complete="off"></el-input>
            </el-form-item>
            <el-form-item label="联系电话" :label-width="formLabelWidth">
              <el-input style="width: auto" v-model="Listform.phone" auto-complete="off"></el-input>
            </el-form-item>
            <el-form-item label="邮件" :label-width="formLabelWidth">
              <el-input style="width: auto" v-model="Listform.email" auto-complete="off"></el-input>
            </el-form-item>
          </el-form>
          <div slot="footer" class="dialog-footer">
            <el-button @click="closeDialog(1)">取 消</el-button>
            <el-button type="primary" @click="insertContact(Listform)" >确 定</el-button>
          </div>
        </el-dialog>
      </div>
      <table id="selfContactList">
        <tr>
          <th>联系人 </th>
          <th>性别 </th>
          <th>职位 </th>
          <th>行政职位 </th>
          <th>部门(学院)</th>
          <th>联系电话 </th>
          <th>邮件 </th>
          <th v-if="kind==1">操作 </th>
        </tr>
        <tr v-for=" (member,key) in contactList" >
          <td>{{member.name}}</td>
          <td v-if="member.sex==1">男</td>
          <td v-if="member.sex==2">女</td>
          <td v-if="member.sex==' '">男</td>
          <td>{{member.position}}</td>
          <td>{{member.zhiwei}}</td>
<!--          <td v-if="member.zhiwei==1">校长</td>-->
<!--          <td v-if="member.zhiwei==2">副校长</td>-->
<!--          <td v-if="member.zhiwei==3">校书记</td>-->
<!--          <td v-if="member.zhiwei==4">院长</td>-->
<!--          <td v-if="member.zhiwei==5">副院长</td>-->
<!--          <td v-if="member.zhiwei==6">院书记</td>-->
<!--          <td v-if="member.zhiwei==7">辅导员</td>-->
          <td>{{member.department}}</td>
          <td>{{member.phone}}</td>
          <td>{{member.email}}</td>
          <td v-if="kind==1">
            <span class="btn2" v-on:click="changeContact(member.id)"><i class="el-icon-edit"></i>修改</span>
            <el-dialog title="修改":visible.sync="dialogflag2" >
              <el-form :model="Listform">
                <el-form-item label="姓名" :label-width="formLabelWidth">
                  <el-input style="width: auto"  v-model="Listform.name" auto-complete="off" ></el-input>
                </el-form-item>
                <el-form-item label="性别" :label-width="formLabelWidth">
                  <select v-model="Listform.sex" style="font-size: 15px;width: auto;height:auto ">
                    <option value="1">男</option>
                    <option value="2">女</option>
                  </select>
                </el-form-item>
                <el-form-item label="身份" :label-width="formLabelWidth">
                  <select v-model="Listform.kind" style="font-size: 15px;width: auto;height:auto ">
                    <option value="1">管理员</option>
                    <option value="2">学生</option>
                  </select>
                </el-form-item>
                <el-form-item label="职位" :label-width="formLabelWidth">
                  <el-input style="width: auto" v-model="Listform.position" auto-complete="off"></el-input>
                </el-form-item>
                <el-form-item label="部门(学院)" :label-width="formLabelWidth">
                  <el-input style="width: auto" v-model="Listform.department" auto-complete="off"></el-input>
                </el-form-item>
                <el-form-item label="联系电话" :label-width="formLabelWidth">
                  <el-input style="width: auto" v-model="Listform.phone" auto-complete="off"></el-input>
                </el-form-item>
                <el-form-item label="邮件" :label-width="formLabelWidth">
                  <el-input style="width: auto" v-model="Listform.email" auto-complete="off"></el-input>
                </el-form-item>
              </el-form>
              <div slot="footer" class="dialog-footer">
                <el-button @click="closeDialog(2)">取 消</el-button>
                <el-button type="primary" @click="updateContact(Listform)" >确 定</el-button>
              </div>
            </el-dialog>
            <span class="btn2" v-on:click="deleteContact(member.id)"><i class="el-icon-delete"></i>删除</span>
          </td>
        </tr>
      </table>
    </div>
  </div>
</template>

<script>

    export default {
        name: 'Mycontact',
        data () {
            return {
                kind:this.$store.getters.getUserkind,
                self1: true,
                inputName: '',
                contactList: [],
                currentpage: 1,
                totalpage: '',
                dialogflag1: false, // 添加
                dialogflag2:false,
                formLabelWidth: '120px',
                Listform:{
                    id:'',
                    name:'',
                    sex:'',
                    kind:'',
                    position:'',
                    department:'',
                    phone:'',
                    email:''
                },
            }
        },
        methods: {
            getContactData () {
                // 获取通讯录数据
                let that = this;
                this.$http.get('/document/contact/getdata?page=1').then(function (res) {
                    console.log(res.data.data)
                    that.contactList = res.data.data[0];
                    that.totalpage = res.data.data[1];
                })
            },
            searchContact: function () {
                // 通讯录搜索
                this.$http.post('/document/contact/search', {
                    contactName: this.inputName, page: this.currentpage
                }).then(res => {
                    console.log(res.data.data)
                    this.contactList = res.data.data[0];
                    this.totalpage = res.data.data[1];
                }).catch(function (error) {
                    console.log(error)
                })
            },
            insertContact: function (Listfrom) {
                // 添加通讯录
                let that = this;
                if(that.Listform.id!=''&& that.Listform.name!=''&&that.Listform.sex!=''&&that.Listform.phone!=''){
                    this.$http.post('/document/contact/insert',{
                        id:that.Listform.id,
                        name: that.Listform.name,
                        sex: that.Listform.sex,
                        kind:that.Listform.kind,
                        phone: that.Listform.phone,
                        department: that.Listform.department,
                        position: that.Listform.position,
                        email: that.Listform.email,
                    }).then(function (res){
                        console.log(res);
                        that.getContactData();
                    })
                }else{
                    alert("添加的信息不能为空！");
                }
                that.dialogflag1=false;
            },
            deleteContact:function (id) {
                // 删除
                console.log(id);
                let that = this;
                this.$http.post('/document/contact/delete', {id: id}).then(function (res) {
                    console.log((res.data))
                    if (res.data.message == 1) {
                        that.getContactData();
                        alert("删除成功！")
                    }
                })
            },
            changeContact:function (id) {
                // 修改
                this.id= id;
                console.log(id);
                let that =this;
                this.$http.post('/document/contact/change', {id: id}).then(function (res) {
                    console.log(res.data);
                    that.Listform=res.data.data;
                    that.dialogflag2=true;
                })
            },
            closeDialog:function (flag) {
                //取消
                if(flag==1){
                    this.dialogflag1=false;
                }
                else if(flag==2){
                    this.dialogflag2=false;
                }

            },
            updateContact:function(Listform){
                console.log(Listform);
                let that = this;
                if(that.Listform.id!=''){
                    this.$http.post('/document/contact/update',{
                        id:Listform.id,
                        name: Listform.name,
                        sex: Listform.sex,
                        kind:Listform.kind,
                        phone: Listform.phone,
                        department: Listform.department,
                        position: Listform.position,
                        email: Listform.email,
                    }).then(function (res) {
                        if(res.data.data==1){
                            alert("修改成功！");
                            that.getContactData();
                            that.dialogflag2=false;
                        }else{
                            alert("修改失败！")
                        }
                    }).catch(function (err) {
                        console.log(err);
                    });
                }else{
                    alert("修改失败！");
                }
            },
        },
        created () {
            // this.kind=this.$store.getters.getUserkind;
            this.getContactData();
            console.log(this.kind);
        }
    }
</script>

<style scoped>
  @import "../../common/css/oa.css";
</style>
