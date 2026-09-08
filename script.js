const word = document.getElementById("word");                                              // HTML里找"word"这个id
const taskList = document.getElementById("task-list");                                     // HTML里找"task-list"这个id
const add = document.getElementById("add-button");                                         // HTML里找"add-button"这个id
const reset = document.getElementById("reset-button");                                     // HTML里找"reset-button"这个id
const importBtn = document.getElementById("import-button");                                // HTML里找"import-button"这个id
const exportBtn = document.getElementById("export-button");                                // HTML里找"export-button"这个id
const toast = document.getElementById("toast");                                            // HTML里找"toast"这个id





importBtn.addEventListener("click", excelImport);                                          //监听 importBtn,当点击时执行函数 excelImport
exportBtn.addEventListener("click", excelExport);                                          //监听 exportBtn,当点击时执行函数 excelExport





/*
Export to Excel
*/
function excelExport() {                                                                   //创建 excelExport 函数

    const excelTasks = [];                                                                 // excelTasks = 空数组

    const tasks = taskList.querySelectorAll(".task-item");                                 // tasks = 选择 taskList 里全部有 class = "task-item" 的

    tasks.forEach(function (item) {                                                        //每个 tasks 各个执行一次,并暂时命名为 item

        const text = item.querySelector("span").textContent;                               // text = 选择 item 的 span ,要里面的文字
        const status = item.querySelector("input").checked;                                // status = 选择 item 的 input ,要看 checkbox 是否打钩的状态

        excelTasks.push({                                                                  //把以下资料加进 excelTasks = 空的数组里
            Task: text,                                                                    // Task: 上面的 text ,也就是 task 的文字
            Completed: status                                                              // Completed: 上面的 status ,也就是 checkbox 是否打钩
        });

    });

    const worksheet = XLSX.utils.json_to_sheet(excelTasks);                                //把 excelTasks 数组里的资料转换成 Excel 的工作表

    const workbook = XLSX.utils.book_new();                                                //创立一个新的 Excel 工作簿

    XLSX.utils.book_append_sheet(workbook, worksheet, "Tasks");                            //把 worksheet 工作表加入 workbook 工作簿,并命名为 "Tasks"

    XLSX.writeFile(workbook, "Todo-List-Backup.xlsx");                                     //把 workbook 保存成 Todo-List-Backup.xlsx 文件并下载
};





/*
Import from Excel
*/
function excelImport() {                                                                   //创建 excelImport 函数
    const file = document.getElementById("excel-file").files[0];                           //从文件选择器里拿用户选择的第一个文件

    if (!file) return;                                                                     //如果没有 file ,终止函数并返回函数的值

    const reader = new FileReader();                                                       //建立 FileReader ,用来读取用户选择的文件 (算是建立一个「读取文件的人」)

    reader.onload = function (event) {                                                     //文件读取完毕后,执行里面的函数,event里面有这次读取文件的资料 (告诉 reader「读完以后要做什么」)

        const excelData = new Uint8Array(event.target.result);                             //把读取到的文件资料转成 Uint8Array 格式 (因为后面的 XLSX.read() 要拿这个资料来解析 Excel )

        const workbook = XLSX.read(excelData, {                                            //用 XLSX 读取刚刚转换后的资料,并得到 Excel 的工作簿 workbook 
            type: "array"
        });

        const excelSheet = workbook.Sheets[workbook.SheetNames[0]];                        //从整本 Excel 里拿第一个工作表

        const excelTasks = XLSX.utils.sheet_to_json(excelSheet);                           //把工作表里的资料转成 Array ,再把每一笔资料变成 Object

        excelTasks.forEach(function (task) {                                               //把 excelTasks 里面的每一笔资料拿出来,暂时命名为 task
            createTask(task.Task, task.Completed);                                         // task 是暂时命名, Task 和 Completed 是 Object 里的资料
            /* 
            也就是excelExport()这里的
            excelTasks.push({
                Task: text,
                Completed: status
                });
            */
        });
        saveTasks();                                                                       //执行 saveTasks()
    };

    reader.readAsArrayBuffer(file);                                                        //开始读取 file ,并把文件读取成 ArrayBuffer (真正开始读取 Excel )
};





/*
功能: 两个都是添加任务
一个是点击触发
一个是 Enter 触发
*/
add.addEventListener("click", function () {                                                //监听 add ,当点击时添加任务
    if (word.value === "") return;                                                         //如果 word 里的内容( value )非常确定是""(空内容),返回,立即停下并退出当前函数

    showToast("Add Task Complete !");                                                      //执行 showToast() + 文字"Add Task Complete !"
    createTask();                                                                          //执行 createTask()
    saveTasks();                                                                           //执行 saveTasks()

    word.value = "";                                                                       //清空 word (也就是 id = word )的内容
    word.focus();                                                                          // word (也就是 id = word ),可以继续输入下一件事,不用点击

});

word.addEventListener("keydown", function (event) {                                        //监听 word ,当按下 Enter 建时添加任务
    if (event.key === "Enter") {                                                           //如果 event.key 非常确定是 Enter 键,继续执行以下函数
        if (word.value === "") return;                                                     //如果 word 里的内容( value )非常确定是""(空内容),返回,立即停下并退出当前函数

        showToast("Add Task Complete !");                                                  //执行 showToast() + 文字"Add Task Complete !"
        createTask();                                                                      //执行 createTask()
        saveTasks();                                                                       //执行 saveTasks()

        word.value = "";                                                                   //清空 word (也就是 id = word )的内容
        word.focus();                                                                      // word (也就是 id = word ),可以继续输入下一件事,不用点击
    }
});





/*
功能: 移除框里的文字
*/
reset.addEventListener("click", removeText);                                               //监听 reset ,当点击 reset 时,执行 removeText()函数

function removeText() {                                                                    //创建函数,命名为 removeText()
    word.value = "";                                                                       //把 word 的内容( value )变成""(空内容)
    word.focus();                                                                          //当 word 获得焦点时,用户可以直接继续在里面打字,无需再点击 (简单理解: 按 Tab 时会跳到输入框,这就是 focus )
};





/*
功能: 创建 task 的 Element ,然后放进 task-list 里
*/
function createTask(taskText = word.value, isCompleted = false) {                          //可以给我两个资料,有资料则替换,没资料默认值是 word.value 和 false
    const task = document.createElement("div");                                            // HTML 里建立空的 <div></div>
    task.classList.add("task-item");                                                       //刚建立空的 <div></div> 再加上 class = "task-item"

    const checkbox = document.createElement("input");                                      // HTML里建立空的 <input>   
    checkbox.type = "checkbox";                                                            // input 的类型是 checkbox
    checkbox.checked = isCompleted;                                                        //根据 isCompleted 来判断 checkbox 是否打钩 (跟if (isCompleted) { 联动的)

    const text = document.createElement("span");                                           // HTML 里建立空的 <span></span>
    text.textContent = taskText;                                                           // span 里面的内容是 taskText ,有资料跟资料, 没有就跟 input 输入的内容

    const del = document.createElement("button");                                          // HTML 里建立空的 <button></button>
    del.textContent = "删除";                                                                // button 里添加文字
    del.classList.add("delete-button");                                                    //刚建立空的 <button></button> 再加上 class = "delete-button"

    if (isCompleted) {                                                                     //如果 isCompleted 是 true
        text.classList.add("completed");                                                   //就在 text (也就是 span )里添加 class = "completed"
    };

    checkbox.addEventListener("change", function () {                                      //当 checkbox 的状态改变时, 执行函数
        if (checkbox.checked) {                                                            //当 checkbox 打钩时
            text.classList.add("completed");                                               //在 text (也就是 span )里添加 class = "completed"
        } else {                                                                           //反之
            text.classList.remove("completed");                                            //在 text 里移除 class = "completed"
        }

        saveTasks();                                                                       //保存到 localStorage
    });

    del.addEventListener("click", function () {                                            //当点击 button 时,执行函数
        task.remove();                                                                     //移除当前 task (也就是 div )
        saveTasks();                                                                       //保存到 localStorage
    });

    task.appendChild(checkbox);                                                            //这时变成 <div class = "task-item"> <input type = "checkbox"> </div>
    task.appendChild(text);                                                                //这时变成 <div class = "task-item"> <input type = "checkbox"> <span> text </span> </div>
    task.appendChild(del);                                                                 //这时变成 <div class = "task-item"> <input type = "checkbox"> <span> text </span> <button class = "delete-button"> 删除 </button> </div>
    taskList.appendChild(task);                                                            //最后,把刚刚整理好的 <div class = "task-item"> <input type = "checkbox"> <span> text </span> <button class = "delete-button"> 删除 </button> </div>, 放进去 <div id = "task-list"> 这里 </div> 里面

};





/*
功能: 暂时保存资料在网页,避免刷新时消失
*/
function saveTasks() {

    const saveTaskArr = [];                                                                //创建一个空数组

    const items = taskList.querySelectorAll(".task-item");                                 //选择 <div id = "task-list"> </div>里面有 class = "task-item"的

    items.forEach(function (item) {                                                        //依序读取 items 里的每一个 task-item , 并暂时存在 item 这个变量名

        const text = item.querySelector("span").textContent;                               // text = 在 item 里找到 span 里面的文字

        const status = item.querySelector("input").checked;                                // status = 在 item 里找到 input 是否打钩

        saveTaskArr.push({                                                                 //按照以下格式, 把资料存进上面的 const saveTaskArr 里
            task: text,                                                                    // task = 上面的 const = text
            completed: status                                                              // completed = 上面的 const status
        });

    });

    const data = JSON.stringify(saveTaskArr);                                              // data = 把 saveTaskArr 转换成字串 string

    localStorage.setItem("tasksBackup", data);                                             //在 localStorage 里储存进名字为 tasksBackup , 然后把刚刚转换好的 data 存进去
};





/*
功能: 将保存在网页的资料调用出来
*/
function loadTasks() {
    const webData = localStorage.getItem("tasksBackup");                                   //从 localStorage 里找名字为 tasksBackup 的资料 (目前还是字串)
    if (!webData) return;                                                                  //如果没有 webData ,终止函数并返回函数的值
    const tasks = JSON.parse(webData);                                                     //把刚刚得到的字串转回 Array

    tasks.forEach(function (task) {                                                        //把 tasks 里面的每一笔资料拿出来, 暂时命名为 task
        createTask(task.task, task.completed);                                             // task 是暂时命名, task 和 completed 是 Object 里的资料
        /* 
           也就是saveTasks()这里的
           saveTaskArr.push({                                                           
            task: word,                                                         
            completed: status                                               
        });
        */
    });

};
loadTasks();                                                                               //执行 loadTasks()





/*
添加新任务的时候跳通知
*/
function showToast(message) {                                                              //建立 showToast 函数, message 是要显示的提示内容
    toast.textContent = message;                                                           //把 message 的内容放进 toast 里面
    toast.style.opacity = "1";                                                             //把 toast 的透明度设置为 1 , 让 toast 显示出来

    setTimeout(function () {                                                               //设置一个计时器,时间到了自动执行里面的函数
        toast.style.opacity = "0";                                                         //把 toast 的透明度设置为 0 ,让 toast 隐藏起来
    }, 2000);                                                                              // 2000 毫秒(也就是 2 秒)后执行
};




