<template>
<!--    个人文件-->
  <div class="display1">
    <button class="btn1" v-on:click="upload" v-bind:class="{active:isActiveOne}">上传文件</button>
    <input id="imFile" style="display: none" type="file" @change="importfxx(this,1)"
    accept="*/*">
    <button class="btn1" v-on:click="newfolder" v-bind:class="{active:isActiveTwo}">新建文件夹</button>

    <div class="display2">
      <table>
        <tr>
          <th>文件名称</th>
          <th>文件大小</th>
          <th>上传时间</th>
          <th>操作</th>
        </tr>
        <tr>
          <th>test</th>
          <th>1</th>
          <th>2019.10</th>
          <th>删除</th>
        </tr>
      </table>
    </div>
  </div>
</template>

<script>
    export default {
        name: "MyDocument",
        data(){
            return{
                isActiveTwo:false,
                isActiveOne:true,
            }
        },
        methods:{
            upload:function () {
                //上传文件
                this.imFile.click()
            },
            newfolder:function () {
                //新建文件夹

            },
            uploadFile: function () { // 点击导入按钮
                this.imFile.click()
            },
            importfxx(obj, flag) {
                let _this = this
                let inputDOM = this.$refs.inputer   // 通过DOM取文件数据
                this.file = event.currentTarget.files[0]
                var rABS = false // 是否将文件读取为二进制字符串
                var f = this.file
                var reader = new FileReader()
                // if (!FileReader.prototype.readAsBinaryString) {
                FileReader.prototype.readAsBinaryString = function (f) {
                    var binary = ''
                    var rABS = false // 是否将文件读取为二进制字符串
                    var pt = this
                    var wb // 读取完成的数据
                    var outdata
                    var reader = new FileReader()
                    reader.onload = function (e) {
                        var bytes = new Uint8Array(reader.result)
                        var length = bytes.byteLength
                        for (var i = 0; i < length; i++) {
                            binary += String.fromCharCode(bytes[i])
                        }
                        var XLSX = require('xlsx')
                        if (rABS) {
                            wb = XLSX.read(btoa(fixdata(binary)), { // 手动转化
                                type: 'base64'
                            })
                        } else {
                            wb = XLSX.read(binary, {
                                type: 'binary'
                            })
                        }
                        // outdata就是你想要的东西 excel导入的数据
                        outdata = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]])
                        // excel 数据再处理
                        let arr = []
                        console.log(flag)
                        if (flag == 1) {
                            outdata.map(v => {
                                let obj = {}
                                obj.job_num = v.工号
                                obj.major = v.专业名称
                                obj.teacherName = v.教师姓名
                                obj.identity = v.身份证号
                                obj.sex = v.性别
                                obj.rank = v.职称
                                obj.post = v.职务
                                obj.contactPhone = v.联系电话
                                obj.email = v.邮箱
                                arr.push(obj)
                            })
                        } else {
                            outdata.map(v => {
                                let obj = {}
                                obj.sno = v.学号
                                obj.major = v.专业
                                obj.stuName = v.姓名
                                obj.grade = v.年级
                                obj.class = v.班级
                                obj.sex = v.性别
                                obj.bornDate = v.出生日期
                                obj.identity = v.身份证号
                                obj.contactPhone = v.联系电话
                                obj.email = v.邮箱
                                arr.push(obj)
                            })
                        }

                        _this.memberList = [...arr]
                        let data = {
                            data: JSON.stringify(_this.memberList)
                        }
                        if (flag == 1) {
                            _this.$http.post('', data).then(body => {
                                console.log(body)
                                alert("导入成功")
                                _this.getTchData()
                            })
                        } else if (flag == 2) {
                            ///缺少路径
                            _this.$http.post('', data).then(body => {
                                alert(body.data.message)
                                _this.getStuData()
                            })
                        }
                    }
                    reader.readAsArrayBuffer(f)
                }
                if (rABS) {
                    reader.readAsArrayBuffer(f)
                } else {
                    reader.readAsBinaryString(f)
                }
            },
        },
        mounted() {
            this.imFile=document.getElementById('imFile')
        }
    }
</script>

<style scoped>
  .display1 {
    padding-left: 5px;
    padding-top: 10px;
  }
  .btn1 {
    font-size: 16px;
    padding: 10px 10px;
    border-top-left-radius: 3px;
    border-top-right-radius: 3px;
    border: 1px solid #ccc;
    cursor: pointer;
    background-color: #fff;
    margin-bottom: -1px;
    color: #000;
    width: 150px;
  }
  .btn1:hover {
    background: #e0e0e0;
  }
  .btn2 {
    width: 80px;
    padding: 7px;
    font-size: 14px;
    border-radius: 3px;
    border: none;
    color: white;
    background-color: #338FFC;
    float: left;
    margin-left: 5px;
    margin-top: 17px;
    margin-bottom: 5px;
  }

  .btn2:hover {
    background-color: #5FA7FE;
  }
  .display2 {
    border: solid 1px #e0e0e0;
    height: 600px;
    width: 100%;
    padding-left: 5px;
    padding-right: 5px;
    background-color: #fff;
  }
  table {
    border-collapse: collapse;
    width: 100%;
    margin-top: 10px;
  }
  th {
    font-size: 14px;
    border: solid 1px #ccc;
    font-weight: bold;
    padding: 5px;
    background-color: #F1F1F1;
    text-align: center;
  }
  table, td {
    border: solid 1px #ccc;
    padding: 5px;
    text-align: center;
    font-size: 14px;
  }
</style>
