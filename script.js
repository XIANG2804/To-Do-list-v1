//建议ExportExcel改成excelExport,看开头就知道负责什么了
//建议把Excel和saveTasks里的text Text completed Completed 分类,要不然回来看还要捋多一遍
//注解想办法位移去同个地方

const word = document.getElementById("word");                               //HTML里找"word"这个id
const taskList = document.getElementById("task_list");                      //HTML里找"task_list"这个id
const add = document.getElementById("addButton");                           //HTML里找"addButton"这个id
const reset = document.getElementById("resetButton");                       //HTML里找"resetButton"这个id
const importBtn = document.getElementById("importButton");//HTML里找"importButton"这个id
const exportBtn = document.getElementById("exportButton");//HTML里找"exportButton"这个id
const toast = document.getElementById("toast");//HTML里找"toast"这个id




importBtn.addEventListener("click", importExcel);//监听importBtn,当点击时执行函数importExcel
exportBtn.addEventListener("click", exportExcel);//监听exportBtn,当点击时执行函数exportExcel






/*
Import from Excel
*/
function importExcel() {//创建importExcel函数
    const file = document.getElementById("excelFile").files[0]; //从文件选择器里拿用户选择的第一个文件

    if (!file) return;     //如果没有file,终止函数并返回函数的值

    const reader = new FileReader();    //建立FileReader,用来读取用户选择的文件 (算是建立一个「读取文件的人」)

    reader.onload = function (event) {  //文件读取完毕后,执行里面的函数,event里面有这次读取文件的资料 (告诉 reader「读完以后要做什么」)

        const data = new Uint8Array(event.target.result); //把读取到的文件资料转成Uint8Array格式 (因为后面的 XLSX.read() 要拿这个资料来解析 Excel)

        const workbook = XLSX.read(data, {  //用XLSX读取刚刚转换后的资料,并得到Excel的工作簿(workbook)
            type: "array"
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]]; //从整本Excel里拿第一个工作表

        const tasks = XLSX.utils.sheet_to_json(sheet); //把工作表里的资料转成Array,再把每一笔资料变成Object

        tasks.forEach(function (task) { //把tasks里面的每一笔资料拿出来,暂时命名为task
            createTask(task.Task, task.Completed);  //task是暂时命名,Task和Completed是Object里的资料
            /* 
            也就是ExportExcel()这里的
            tasks.push({
                Task: text,
                Completed: completed
                });
                */
        });
        saveTasks();                                                            //执行saveTasks()
    };

    reader.readAsArrayBuffer(file); //开始读取file,并把文件读取成ArrayBuffer (真正开始读取 Excel)
};





//Export to Excel
function exportExcel() {//创建exportExcel函数

    const tasks = [];   //tasks = 空数组

    const items = taskList.querySelectorAll(".task-item");  //items = 选择taskList里全部有class = "task-item"的

    items.forEach(function (item) { //每个items各个执行一次,并暂时命名为item

        const text = item.querySelector("span").textContent; //text = 选择item的span,要里面的文字
        const completed = item.querySelector("input").checked;  //completed = 选择item的input,要看checkbox是否打钩的状态

        tasks.push({    //把以下资料加进 tasks = 空的数组里
            Task: text, //Task: 上面的text,也就是task的文字
            Completed: completed    //Completed: 上面的completed,也就是checkbox是否打钩
        });

    });

    const worksheet = XLSX.utils.json_to_sheet(tasks); //把tasks数组里的资料转换成Excel的工作表

    const workbook = XLSX.utils.book_new(); //创立一个新的Excel工作簿

    XLSX.utils.book_append_sheet(workbook, worksheet, "Tasks"); //把worksheet工作表加入workbook工作簿,并命名为"Tasks"

    XLSX.writeFile(workbook, "todo-list.xlsx"); //把workbook保存成todo-list.xlsx文件并下载
};





/*
功能: 两个都是添加任务
一个是点击触发
一个是Enter触发
*/
add.addEventListener("click", function () {                                 //监听add,当点击时添加任务
    if (word.value === "") return;                                          //如果word里的内容(value)非常确定是""(空内容),返回,立即停下并退出当前函数

    showToast("Add Task Complete !");                                       //执行showToast() + 文字"Add Task Complete !"
    createTask();                                                              //执行createTask()
    saveTasks();                                                            //执行saveTasks()

    word.value = "";                                                        //清空word(也就是id=word)的内容
    word.focus();                                                           //word(也就是id=word),可以继续输入下一件事,不用点击

});

word.addEventListener("keydown", function (event) {                         //监听word,当按下Enter建时添加任务
    if (event.key === "Enter") {                                            //如果event.key非常确定是Enter建,继续执行以下函数
        if (word.value === "") return;//如果word里的内容(value)非常确定是""(空内容),返回,立即停下并退出当前函数

        showToast("Add Task Complete !");//执行showToast() + 文字"Add Task Complete !"
        createTask();//执行createTask()
        saveTasks();//执行saveTasks()

        word.value = "";                                                        //清空word(也就是id=word)的内容
        word.focus();                                                           //word(也就是id=word),可以继续输入下一件事,不用点击
    }
});





/*
功能: 移除框里的文字
*/
reset.addEventListener("click", removeText);                                //监听reset,当点击reset时,执行removeText()函数

function removeText() {                                                     //创建函数,命名为removeText()
    word.value = "";                                                        //把word的内容(value)变成""(空内容)
    word.focus();                                                           //当word获得焦点时,用户可以直接继续在里面打字,无需再点击 (简单理解: 按Tab时会跳到输入框,这就是focus)
};





/*
功能: 创建task的Element,然后放进task_list里
*/
function createTask(taskText = word.value, isCompleted = false) {           //可以给我两个资料,有资料则替换,没资料默认值是word.value和false
    const task = document.createElement("div");                             //HTML里建立空的<div></div>
    task.classList.add("task-item");                                        //刚建立空的<div></div>再加上class="task-item"

    const checkbox = document.createElement("input");                       //HTML里建立空的<input>   
    checkbox.type = "checkbox";                                             //input的类型是checkbox
    checkbox.checked = isCompleted;                                         //根据isCompleted来判断checkbox是否打钩 (跟if (isCompleted) { 联动的)

    const text = document.createElement("span");                            //HTML里建立空的<span></span>
    text.textContent = taskText;                                            //span里面的内容是taskText,有资料跟资料, 没有就跟input输入的内容

    const del = document.createElement("button");                           //HTML里建立空的<button></button>
    del.textContent = "删除";                                               //button里添加文字
    del.classList.add("delete-button");                                     //刚建立空的<button></button>再加上class="delete-button"

    if (isCompleted) {          //如果isCompleted是true
        text.classList.add("completed");    //就在text(也就是span)里添加class="completed"
    };

    checkbox.addEventListener("change", function () {                        //当checkbox的状态改变时, 执行函数
        if (checkbox.checked) {                                              //当checkbox打钩时
            text.classList.add("completed");                                //在text(也就是span)里添加class="completed"
        } else {                                                            //反之
            text.classList.remove("completed");                             //在text里移除class="completed"
        }

        saveTasks();                                                        //保存到localStorage
    });

    del.addEventListener("click", function () {                              //当点击button时,执行函数
        task.remove();                                                      //移除当前task(也就是div)
        saveTasks();                                                        //保存到localStorage
    });

    task.appendChild(checkbox);//这时变成<div class="task-item"> <input type="checkbox"> </div>
    task.appendChild(text);//这时变成<div class="task-item"> <input type="checkbox"> <span>text</span> </div>
    task.appendChild(del);//这时变成<div class="task-item"> <input type="checkbox"> <span>text</span> <button class="delete-button">删除</button> </div>
    taskList.appendChild(task);//最后,把刚刚整理好的<div class="task-item"> <input type="checkbox"> <span>text</span> <button class="delete-button">删除</button> </div>, 放进去<div id="task_list">这里</div>里面

};





/*
功能: 暂时保存资料在网页,避免刷新时消失
*/
function saveTasks() {

    const arr = [];                                             //创建一个空数组

    const items = taskList.querySelectorAll(".task-item");      //选择<div id="task_list"></div>里面有class="task-item"的

    items.forEach(function (item) {                             //依序读取items里的每一个task-item, 并暂时存在item这个变量名

        const text = item.querySelector("span").textContent;    //text = 在item里找到span里面的文字

        const completed = item.querySelector("input").checked;  //completed = 在item里找到input是否打钩

        arr.push({                                              //按照以下格式, 把资料存进上面的const arr里
            text: text,                                         //text = 上面的const = text
            completed: completed                                //completed = 上面的const completed
        });

    });

    const data = JSON.stringify(arr);                           //data = 把arr转换成字串string

    localStorage.setItem("tasks", data);                        //在localStorage里储存进名字为tasks, 然后把刚刚转换好的data存进去
};





/*
功能: 将保存在网页的资料调用出来
*/
function loadTasks() {
    const data = localStorage.getItem("tasks");             //从localStorage里找名字为tasks的资料 (目前还是字串)
    if (!data) return;                                      //如果没有data,终止函数并返回函数的值
    const tasks = JSON.parse(data);                         //把刚刚得到的字串转回Array

    tasks.forEach(function (task) {                         //把tasks里面的每一笔资料拿出来, 暂时命名为task
        createTask(task.text, task.completed);              //task是暂时命名,text和completed是Object里的资料
    });

};
loadTasks();                                                //执行loadTasks()





/*
添加新任务的时候跳通知
*/
function showToast(message) {   //建立showToast函数,message是要显示的提示内容
    toast.textContent = message;    //把message的内容放进toast里面
    toast.style.opacity = "1";      //把toast的透明度设置为1, 让toast显示出来

    setTimeout(function () {        //设置一个计时器,时间到了自动执行里面的函数
        toast.style.opacity = "0";  //把toast的透明度设置为0,让toast隐藏起来
    }, 2000);                       //2000毫秒(也就是2秒)后执行
};






































































































