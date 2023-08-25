const docx4js = require("docx4js");

docx4js.load("~/test.docx").then(docx=>{
	//you can render docx to anything (react elements, tree, dom, and etc) by giving a function
	docx.render(function createElement(type,props,children){
		return {type,props,children}
	})

	//or use a event handler for more flexible control
	const ModelHandler=require("docx4js/lib/openxml/docx/model-handler").default
	class MyModelhandler extends ModelHandler{
		onp({type,children,node},  officeDocument){

		}
	}

	const handler=new MyModelhandler()
	handler.on("*",function({type,children,node},  officeDocument){
		console.log("found model:"+type)
	})
	handler.on("r",function({type,children,node},  officeDocument){
		console.log("found a run")
	})

	docx.parse(handler)

	//you can change content on docx.officeDocument.content, and then save
	docx.officeDocument.content("w\\:t").text("hello")
	docx.save("~/changed.docx")

})

//you can create a blank docx
docx4js.create().then(docx=>{
	//do anything you want
	docx.save("~/new.docx")
})
