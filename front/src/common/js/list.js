$(function(){
  $('#NAME').focus().blur(checkName);
  $('#password').blur(checkPassword);
});

$(".file").on("change","input[type='file']",function () {
  var filepath = $(this).val();
  if(filepath.indexOf("jpg")!=-1||filepath.indexOf("png")!=-1)
  {
    $(".fileerrorTip1").html("").hide();
    var arr = filepath.split('\\');
    var fileName = arr[arr.length-1];
    $(".showFileName").html(fileName);
  }else{
    $(".showFileName").html();
    $(".fileerrorTip1").html("未上传文件").show();
    return false
  }
})
