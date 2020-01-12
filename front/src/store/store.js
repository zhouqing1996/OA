import Vue from 'vue'
import Vuex from 'vuex'

Vue.use(Vuex)
const state={
    //全局变量
    userid:JSON.parse(localStorage.getItem("vuex"))?JSON.parse(localStorage.getItem("vuex"))['userid']:"",
    username:JSON.parse(localStorage.getItem("vuex"))?JSON.parse(localStorage.getItem("vuex"))['username']:"",
    userdpt:JSON.parse(localStorage.getItem("vuex"))?JSON.parse(localStorage.getItem("vuex"))['userdpt']:"",
    userkind:JSON.parse(localStorage.getItem("vuex"))?JSON.parse(localStorage.getItem("vuex"))['userkind']:"",
    login:JSON.parse(localStorage.getItem("vuex"))?JSON.parse(localStorage.getItem("vuex"))['login']:false,
  Loading:JSON.parse(localStorage.getItem("vuex"))?JSON.parse(localStorage.getItem("vuex"))['Loading']:false,
  };
const  mutations={
    //定义各种方法，一般用于axios
  setUserid:function (state,userid) {
    state.userid=userid;
    localStorage.setItem("vuex",JSON.stringify(state));
  },
  delUserid:function (state) {
    state.userid=" ";
    localStorage.setItem("vuex",JSON.stringify(state));
  },
  setUsername:function (state,username) {
    state.username=username;
    localStorage.setItem("vuex",JSON.stringify(state));
  },
  delUsername:function (state) {
    state.username=" ";
    localStorage.setItem("vuex",JSON.stringify(state));
  },
  setUserdpt:function (state,userdpt) {
    state.userdpt=userdpt;
    localStorage.setItem("vuex",JSON.stringify(state));
  },
  delUserdpt:function (state) {
    state.userdpt=" ";
    localStorage.setItem("vuex",JSON.stringify(state));
  },
  setUserkind:function (state,userkind) {
    state.userkind=userkind;
    localStorage.setItem("vuex",JSON.stringify(state));
  },
  delUserkind:function (state) {
    state.userkind=" ";
    localStorage.setItem("vuex",JSON.stringify(state));
  },
  showByLogin:function(state){
    state.login=true
    localStorage.setItem("vuex",JSON.stringify(state));
    },
  hideByLogin:function(state){
    state.login=false
    localStorage.setItem("vuex",JSON.stringify(state));
    },

  showLoading:function(state){
    state.Loading = true
    localStorage.setItem("vuex",JSON.stringify(state));
  },
  hideLoading:function (state) {
    state.Loading = false
    localStorage.setItem("vuex",JSON.stringify(state));

  }
};
const  getters={
    getUserid:(state)=>{
      return state.userid
    },
    getUsername:(state)=>{
      return state.username
    },
    getUserdpt:(state)=>{
      return state.userdpt
    },
  getUserkind:(state)=>{
      return state.userkind
  },
  getLogin:(state)=>{
      return state.login
  },
  getLoading:(state)=>{
      return state.Loading
  }
  };
const  actions={
  login:({commit},obj)=> {
    commit('setUserid', obj.id);
    commit('setUsername', obj.name);
    commit('setUserdpt', obj.department);
    commit('setUserkind',obj.kind);
    commit('showByLogin',obj.login);
  },
  logout:({commit})=>{
    commit('delUserid');
    commit('delUsername');
    commit('delUserdpt');
    commit('delUserkind');
    commit('hideByLogin');
  },
  showByLogin:({commit})=>{
    commit('showByLogin')
  },
  hideByLogin:({commit})=>{
    commit('hideByLogin')
  },
  showLoading:({commit})=>{
    commit('showLoading')
  },
  hideLoading:({commit})=>{
    commit('hideLoading')
  }
}
export default new Vuex.Store({
  state,
  mutations,
  getters,
  actions
})
