<html>
<head>

</head>
<body>


<script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	var slider = $('div.slider');
	var slides = slider.children('div.slides').children();
	var display = slider.children('div.display');
	
	var slideNum = 0;
	slides.eq(slideNum).clone().appendTo(display);
	setInterval(function(){
		var nextSlideNum = slideNum + 1;
		var nextSlide = slides.eq(nextSlideNum).clone();
		
		nextSlide.appendTo(display).hide().fadeIn(1500, function(){
			nextSlide.siblings().remove();
		});
		
		slideNum = (slideNum + 1 < slides.length) ? (slideNum + 1) : 0;
	}, 5000);
	
});
</script>
<style type="text/css">
div.slider{
	position: relative;
	height: 359px;
	width: 835px;
}
div.slider div.slides{
	display: none;
}
div.slider div.display{
	position: absolute;
	height: 100%;
	width: 100%;
}
div.slider div.display img{
	position: absolute;
	max-height: 100%;
	max-width: 100%;
}
div.slider div.text{
	background-image: url('img/slider/white.png');
	border: 1px solid #5C3317;
	position: absolute;
	color: #5C3317;
	
	font-family: "Times New Roman";
	font-size: 22px;
	padding: 5px;
}
div.slider div.text.top{
	border-left-width: 0px;
	top: 20px;
}
div.slider div.text.bottom{
	border-right-width: 0px;
	bottom: 20px;
	right: 0px;
}
</style>
<div class="slider">
	<div class="slides"><img src="images/HomePic-FILLER01.jpg" height="338" width="576"/><img src="images/HomePic-FILLER02.png" height="338" width="576"/></div>

	<div class="display">
	</div>

</div>
</body>
</html>