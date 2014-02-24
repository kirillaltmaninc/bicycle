//window.onload = function () {
	/* Example 1 */
	var overlay_1 = new OverlayJS({
		selector: "example_1",
		width: 300,
		height: 240,
		buttons: {
			"Close": function (button) {
				this.close();
			}
		}
	});
	document.getElementById("button_1").onclick = function () {
alert("hi");
		overlay_1.open();
	};
	
//};