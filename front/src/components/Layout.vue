<template>
  <div>
    <Condition v-if="start"> </Condition>
    <div id="wrapper">
      <Header v-show="start" > </Header>
      <div  v-bind:class="{'main':loginflag}">
        <keep-alive>
          <router-view></router-view>
        </keep-alive>
      </div>
      <div class="clearfix"></div>
    </div>
  </div>
</template>

<script>
    import Header from "./Header";
    import '../utils/jquery-slimscroll/jquery.slimscroll'
    import '../common/js/klorofil-common'
    import Condition from "./Condition";
    export default {
        name: "Layout",
        data(){
            return{
                loginflag:"",
            }
        },
        computed:{
            start:function () {
               return this.loginflag=this.$store.getters.getLogin;
            }
        },
        components: {
            Condition, Header}
    }
</script>

<style >
  @import "../common/css/common.css";
  @import "../common/css/reset.css";
  @import '../utils/bootstrap/css/bootstrap.min.css';
  @import "../utils/font-awesome/css/font-awesome.css";
  .login_content{
    width: 100%;
  }

</style>
