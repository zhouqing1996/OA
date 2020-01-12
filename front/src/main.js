// The Vue build version to load with the `import` command
// (runtime-only or standalone) has been set in webpack.base.conf with an alias.
import Vue from 'vue'
import App from './App'
import Layout from "./components/Layout";
import routes from './router/index'
import ElementUI from 'element-ui'
import 'element-ui/lib/theme-chalk/index.css'
import VueRouter from "vue-router";
import VueResource from 'vue-resource'
import axios from "axios";
import Vueaxios from 'vue-axios';
import store from "./store/store";
import uploader from 'vue-simple-uploader';


Vue.use(uploader)
Vue.config.productionTip = false
Vue.use(ElementUI)
Vue.use(VueRouter)
// Vue.use(VueResource)
Vue.use(Vueaxios,axios)
import baseConfig from 'config'
axios.interceptors.request.use(function (config) {
  if(config.url.indexOf("/translate")<0)
    config.url=baseConfig.apiBaseUrl + config.url;

  // store.dispatch('showLoading')

  return config
},function (error) {
  return Promise.reject(error)
})
axios.interceptors.response.use(function (response) {
  // store.dispatch('hideLading')
  return response;
},function (error) {
  return Promise.reject(error)
});
const router=new VueRouter({
  mode:'history',
  scrollBehavior: () => ({ y: 0 }),
  base:"/OA/front",
  routes,
  linkActiveClass:'active'
})
// /* eslint-disable no-new */
new Vue({
  el: '#app',
  router,
  store,
  components: { Layout },
  template: '<Layout/>'
})
