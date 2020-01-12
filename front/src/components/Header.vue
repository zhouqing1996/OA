<template>
  <div>
    <nav class="navbar navbar-default navbar-fixed-top">
      <div class="brand">
        <router-link :to="{path:'/home/index'}">智慧校园</router-link>
      </div>
      <div class="container-fluid">
        <div id="navbar-menu" style="float: right">
          <ul class="nav navbar-nav navbar-right">
            <li class="dropdown">
              <a href="#" class="dropdown-toggle" data-toggle="dropdown">
                <img src="../common/images/user.png" class="img-circle">
                <span>{{getName}}</span>
              </a>
            <li class="dropdown">
              <router-link :to="{path:'/my'}">
                <span>个人中心</span>
              </router-link>
            </li>
            <li class="dropdown">
              <a v-on:click="logout" class="dropdown-toggle" data-toggle="dropdown">
              <span>退出</span>
              </a>
            </li>
          </ul>
        </div>
      </div>
    </nav>
  </div>
</template>

<script>
    export default {
        name: "Header",
        data(){
            return{
                username:"",
            }
        },
        computed:{
          getName:function () {
              return this.$store.getters.getUsername;
          }
        },
        methods:{
            logout:function () {
                this.$store.dispatch('logout');
                alert("退出登录");
                this.$router.push('/home/login');
            }
        }
    }
</script>

<style scoped>

</style>
