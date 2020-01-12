$(function(){
	$('#passuser').focus().blur(checkName);
	$('#passchange').blur(checkPassword);
});

function checkName(){
	var name = $('#passuser').val();
	if(name == null || name == ""){
		//提示错误
		$('#user-msg').html("用户名不能为空");
		return false;
	}
	var reg = /^\w{3,10}$/;
	if(!reg.test(name)){
		$('#user-msg').html("输入3-10个字母或数字或下划线");
		return false;
	}
	$('#user-msg').empty();
	return true;
}

function checkPassword(){
	var password = $('#passchange').val();
	if(password == null || password == ""){
		//提示错误
		$('#pass-msg').html("密码不能为空");
		return false;
	}
	var reg = /^\w{3,10}$/;
	if(!reg.test(password)){
		$('#pass-msg').html("输入3-10个字母或数字或下划线");
		return false;
	}
	$('#pass-msg').empty();
	return true;
}