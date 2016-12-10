﻿<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Niceforms</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="text/javascript" src="js/jquery.js"></script>
<script type="text/javascript" src="js/nicejforms.js"></script>
<style type="text/css" media="screen">@import url(css/niceforms-default.css);</style>
</head>
 
<body>
<div id="container">
<h1 style="margin-bottom:30px;">NiceJForms</h1>
<form action="../../vars.php" method="POST" class="niceform">
	<select size="1" id="mySelect1" name="mySelect1" class="width_320">
		<option selected="selected" value="Test area no.1">Test area no.1</option>
		<option value="Another test option">Another test option</option>
		<option value="And another one">And another one</option>
		<option value="One last option for me">One last option for me</option>
		<option value="This is one really really long option right here just to test it out">This is one really really long option right here just to test it out</option>
	</select>
	<br />
	<select size="1" id="mySelect2" name="mySelect2" class="width_160">
		<option value="Test area no.2">Test area no.2</option>
		<option value="Another test">Another test</option>
		<option selected="selected" value="And another one">And another one</option>
		<option value="And yet another one">And yet another one</option>
		<option value="One last option for me">One last option for me</option>
	</select>
	<br />
	<input type="radio" name="radioSet" id="option1" value="foo" checked="checked" /><label for="option1">foo</label><br />
	<input type="radio" name="radioSet" id="option2" value="bar" /><label for="option2">bar</label><br />
	<input type="radio" name="radioSet" id="option3" value="another option" /><label for="option3">another option</label><br />
	
	<br />
	<input type="checkbox" name="checkSet1" id="check1" value="foo" /><label for="check1">foo</label><br />
	<input type="checkbox" name="checkSet2" id="check2" value="bar" checked="checked" /><label for="check2">bar</label><br />
	<input type="checkbox" name="checkSet3" id="check3" value="another option" /><label for="check3">another option</label><br />
	<br />
	<label for="textinput">Username:</label><br />
	<input type="text" id="textinput" name="textinput" size="12" /><br />
	<label for="passwordinput">Password:</label><br />
	<input type="password" id="passwordinput" name="passwordinput" size="20" /><br />
	<br />
	<label for="textareainput">Comments:</label><br />
	<textarea id="textareainput" name="textareainput" rows="10" cols="30"></textarea><br />
	<br />
	<input type="submit" value="Submit this form" />
</form>
 
<h2 style="margin-top:30px;">NiceJForms v0.1</h2>
<p>&copy; Lucian Lature - <a href="emailto:lucian.lature@gmail.com">email to me</a></p>
<p>Feel free to use and modify but please provide credits.</p>
</div>
 
<script type="text/javascript"> 
$(document).ready(function(){$.NiceJForms.build()});
</script>
 
</body>
</html>