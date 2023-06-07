<style>
@media (min-width: 768px) {
  .modal-xl {
    width: 90%;
   max-width:1200px;
  }
}
div.panel:first-child {
    margin-top:-20px;
}
.TextoNormal {
	font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;
	color:#333;
	font-size:12px;
}
.TextoNormalRojo {
	font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;
	color:#900;
	font-size:12px;
}
.TextoNormalBlanco {
	font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;
	color:#FFF;
	font-size:13px;
}
.TextoNormalAzul {
	font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;
	color:#5fb0e4;
	font-size:12px;
}

div.treeview {
    min-width: 100px;
    min-height: 100px;
    
    max-height: 256px;
    overflow:auto;
	
	padding: 4px;
	
	margin-bottom: 20px;
	
	color: #369;
	
	border: solid 1px;
	border-radius: 4px;
	
}
div.treeview ul:first-child:before {
    display: none;
}
.treeview, .treeview ul {
    margin:0;
    padding:0;
    list-style:none;
    font-size:14px;
	font-family:Verdana, Geneva, sans-serif;
	color: #369;
}
.treeview ul {
    margin-left:1em;
    position:relative
}
.treeview ul ul {
    margin-left:.5em
}
.treeview ul:before {
    content:"";
    display:block;
    width:0;
    position:absolute;
    top:0;
    left:0;
    border-left:1px solid;
    
    /* creates a more theme-ready standard for the bootstrap themes */
    bottom:15px;
}
.treeview li {
    margin:0;
    padding:0 1em;
    line-height:2em;
    /*font-weight:700;*/
    position:relative
}
.treeview ul li:before {
    content:"";
    display:block;
    width:10px;
    height:0;
    border-top:1px solid;
    margin-top:-1px;
    position:absolute;
    top:1em;
    left:0
}
.tree-indicator {
    margin-right:5px;
    
    cursor:pointer;
}
.treeview li a {
    text-decoration: none;
    color:inherit;
    
    cursor:pointer;
}
.treeview li button, .treeview li button:active, .treeview li button:focus {
    text-decoration: none;
    color:inherit;
    border:none;
    background:transparent;
    margin:0px 0px 0px 0px;
    padding:0px 0px 0px 0px;
    outline: 0;
}


.center {
    margin-top:50px;   
}

.btn-hover-green:hover{
	background-color:#b7fac7;
}

.btn-hover-default-red:hover{
	background-color:#fba8b0;
}

<!-- Para el menu -->
.wrapper{
  float:left;
  width:100%;
  min-height:50px;
  margin-top:-100px;
}
.navigation{
    float: left;
    width: 75%;
    text-align: left;
	padding-left:10px;
	  background-color:#343a40;
	  color:#fff;
	  
}

.navigation ul{
    margin: 0;
    padding: 0;
    float: none;
    width: auto;
    list-style: none;
    display: inline-block;

}
.navigation ul li{
    float: left;
    width: auto;
    margin-right: 60px;
    position: relative;
}
.navigation ul li:last-child{
    margin: 0;
}
.navigation ul li a{
    float: left;
    width: 100%;
    color: #fff;
    padding: 18px 0;
    font-size: 12px;
    line-height: normal;
    text-decoration:none;
    box-sizing:border-box;
    text-transform: uppercase;
    font-family: 'Montserrat', sans-serif;      -webkit-transition:color 0.3s ease;
    transition:color 0.3s ease;
}
.navigation .children {
    position: absolute;
    top: 100%;
    z-index: 1000;
    margin: 0;
    padding: 0;
    left: 0;
    min-width: 240px;
    background-color: #fff;
    border: solid 1px #dbdbdb;
    opacity: 0;
    -webkit-transform-origin: 0% 0%;
    transform-origin: 0% 0%;
    -webkit-transition: opacity 0.3s, -webkit-transform 0.3s;
    transition: opacity 0.3s, -webkit-transform 0.3s;
    transition: transform 0.3s, opacity 0.3s;
    transition: transform 0.3s, opacity 0.3s, -webkit-transform 0.3s;
}
.navigation ul li .children  {
    -webkit-transform-style: preserve-3d;
    transform-style: preserve-3d;
    -webkit-transform: rotateX(-75deg);
    transform: rotateX(-75deg);
    visibility: hidden;
}
.navigation ul li:hover > .children  {
    -webkit-transform: rotateX(0deg);
    transform: rotateX(0deg);
    opacity: 1;
    visibility: visible;
}
.navigation ul li .children .children{
	left: 100%;
	top: 0;
}
.navigation ul li.last .children{
	right: 0;
	left: auto;
}
.navigation ul li.last .children .children{
	right: 100%;
	left: auto;
}
.navigation ul li .children li{
	float: left;
	width: 100%;
  margin:0;
}
.navigation ul li .children  a {  
    display: block;
    font-family: "Montserrat", sans-serif;
    text-transform: uppercase;
    font-weight: 700;
    font-size: 11px;
    color: #333;
    text-align: left;
    line-height: 1.5em;
    padding: 16px 30px;
    letter-spacing: normal;
    border-bottom: 1px solid #dbdbdb;
    -webkit-transition: background-color 0.3s ease;
    transition: background-color 0.3s ease;
}
.navigation ul li .children  a:hover{
	color: #fff;
  background-color:#949596;
}
.navigation ul li a:hover{
  color:#fff;
}

/* Tooltip container */
.tooltip1 {
  position: relative;
  display: inline-block;

  cursor:pointer
}

/* Tooltip text */
.tooltip1 .tooltiptext {
  visibility: hidden;
  width: 500px;
  background-color: #555;
  color: #fff;
  text-align: left;
  padding: 5px 5px;
  border-radius: 6px;
  
  /* Para posicion derecha */
  top:-5px;
  left:130%;

  /* Position the tooltip text */
  position: absolute;
  z-index: 1;
  /*bottom: 125%;
  left: 50%;*/
/*  margin-left: -60px;*/

  /* Fade in tooltip */
  opacity: 0;
  transition: opacity 0.3s;
}

/* Tooltip arrow */
.tooltip1 .tooltiptext::after {
  content: "";
    position: absolute;
    top: 15px;
    right: 100%;
    margin-top: -5px;
    border-width: 5px;
    border-style: solid;
    border-color: transparent #555 transparent transparent;
}

/* Show the tooltip text when you mouse over the tooltip container */
.tooltip1:hover .tooltiptext {
  visibility: visible;
  opacity: 1;
}

</style>