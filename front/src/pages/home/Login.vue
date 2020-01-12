<template>
  <div class="page">
    <div class="loginwarrp">
      <div class="logo">智慧校园</div>
      <div class="login_form">
          <li class="login-item">
            <span>用户名：</span>
            <input type="text" id="NAME" name="NAME" class="login_input" v-model="name">
            <span id="count-msg" class="error"></span>
          </li>
          <li class="login-item">
            <span>密　码：</span>
            <input type="password" id="password" name="password" class="login_input" v-model="password">
            <span id="password-msg" class="error"></span>
          </li>
          <div class="clearfix"></div>
          <li class="login-sub">
            <input type="button"  value="登录" v-on:click="getLogin"/>
          </li>
      </div>
    </div>
  </div>
</template>

<script>
  import '../../common/js/login.js'
  export default {
      name: "Login",
      data(){
          return{
              name:"",
              password:"",
      }
    },
    methods: {
        getLogin () {
            this.$http.post('/home/index/login',{ username:this.name,password:this.password })
                .then((res) => {
                    if(res.data.data!=null){
                        console.log(res.data.data);
                        this.$store.dispatch('login',res.data.data);
                        alert(res.data.message);
                        this.$router.push('/home/content');
                    }else{
                        console.log(res.data);
                        alert(res.data.message);
                    }
                }, (err) => {
                    console.log(111,err)
                })
        }
    },
  }

</script>

<style scoped>
  @import "../../common/css/login.css";
</style>
