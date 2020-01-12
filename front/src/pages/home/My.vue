<template>
  <div style="padding-top:10px;">
    <button class="btn1" v-bind:class="{ active:activeTab==2 }" @click="activeTab=2"  >我的消息</button>
    <button class="btn1" v-bind:class="{ active:activeTab==1 }" @click="activeTab=1">密码修改</button>

    <div class="display" v-show="activeTab==1">
      <div class="account">
        <el-form :model="ruleForm2" status-icon :rules="rules2" ref="ruleForm2" label-width="100px" class="demo-ruleForm">
          <el-form-item label="旧密码" prop="pass">
            <el-input type="password" v-model="ruleForm2.old" autocomplete="off"></el-input>
          </el-form-item>
          <el-form-item label="新密码" prop="pass">
            <el-input type="password" v-model="ruleForm2.pass" autocomplete="off"></el-input>
          </el-form-item>
          <el-form-item label="确认密码" prop="checkPass">
            <el-input type="password" v-model="ruleForm2.checkPass" autocomplete="off"></el-input>
          </el-form-item>
          <el-form-item>
            <el-button type="primary" @click="submitForm('ruleForm2')">提交</el-button>
          </el-form-item>
        </el-form>
      </div>
    </div>
  </div>

</template>

<script>
  export default {
    name: 'My',
    data() {
      var validatePass = (rule, value, callback) => {
        if (value === '') {
          callback(new Error('请输入密码'));
        } else {
          if (this.ruleForm2.checkPass !== '') {
            this.$refs.ruleForm2.validateField('checkPass');
          }
          callback();
        }
      };
      var validatePass2 = (rule, value, callback) => {
        if (value === '') {
          callback(new Error('请再次输入密码'));
        } else if (value !== this.ruleForm2.pass) {
          callback(new Error('两次输入密码不一致!'));
        } else {
          callback();
        }
      };
      return {
        activeTab:"2",
        message: [],
        ruleForm2: {
          pass: '',
          checkPass: '',
          old: ''
        },
        rules2: {
          pass: [
            { validator: validatePass, trigger: 'blur' }
          ],
          checkPass: [
            { validator: validatePass2, trigger: 'blur' }
          ],
          old: [
            { validator: validatePass, trigger: 'blur' }
          ]
        },

        totlepage: 1,//总页数
        visiblepage: 5,//可见页数
        logCurrentpage:1,//当前页
      };
    },
    methods: {
      load(){
        let data={
          oldPassword:this.ruleForm2.old,
          newPassword:this.ruleForm2.pass
        }
        this.$http.post('/home/index/alter-account', data)
          .then((res) => {
            console.log(res.data)
            if (res.data.message == "200") {
              alert(res.data.data);
              this.$store.dispatch("logout");
              this.$router.push('/home/login');
            }
            if (res.data.message != "200") {
              alert(res.data.data)
            }
          }, (err) => {
            console.log(err)
          })
      },
      submitForm(formName) {
        this.$refs[formName].validate((valid) => {
          if (valid) {
            this.load()
          } else {
            console.log('error submit!!');
            return false;
          }
        });
      },
      resetForm(formName) {
        this.$refs[formName].resetFields();
      },
      getPrivateLetter() {
        let data = {
          userId: this.$store.getters.getUser
        }
        console.log(data)
        this.$http.post('/home/index/get-message', data)
          .then((res) => {
            if (res.data.data != null) {
              console.log(res.data.data)
              this.message = res.data.data.models;
            }
          }, (err) => {
            console.log(err)
          })
      },
    },
    created() {
      this.getPrivateLetter()
    }
  }
</script>

<style scoped>
  .btn1{
    font-size: 1.3em;
    padding: 6px 6px;
    width:100px;
    border-top-left-radius: 3px;
    border-top-right-radius: 3px;
    border: 1px solid #ccc;
    cursor: pointer;
    background: #f0f0f0;
    color:#000;
  }
  .btn1:hover{
    background: #e0e0e0;
  }
  .active{
    background: #e0e0e0;
  }

  .display{
    border:solid 1px  #e0e0e0;
    height: 600px;
    width: 98%;
    padding-left:10px;
    padding-right:10px;
    background-color:  #fff;
  }
  .meeting{
    margin:10px 0 10px 0;
    font-weight: bold;
    background-color:  #ebe8e8;
    border-radius: 5px;
    padding:5px;
  }
  .sort{
    color:#1C93FC;
  }
  .sort:hover{
    color:#5CB0FA;
  }
  table{
    border-collapse: collapse;
    width:100%;
    margin-top: 10px;
    margin-left: 0px;
  }
  th{
    border:solid 1px #ccc;
    font-weight: bold;
    padding:5px;
    background-color: #F1F1F1;
    text-align: center;
  }
  table,td{
    border:solid 1px #ccc;
    padding:10px 15px;
    text-align: center;
  }
  .page{
    text-align: center;
  }
  .account{
    margin-top: 30px;
    width: 40%;
  }
  @media screen and (max-width: 768px) {
    .account{
      width: 100%;
    }
  }
</style>
