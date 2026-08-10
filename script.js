const word = document.getElementById("word");    /*HTML里找到<input type="text" id="word">这个容器*/
const taskList = document.getElementById("task_list");    /*HTML里找到<div id="task_list"></div>这个容器*/



word.addEventListener("keydown", function (event) {
    if (event.key === "Enter") {
        if (word.value === "") return;

        showToast("Add Task Complete !");
        addTask();
        saveTasks();
    }
});




const add = document.getElementById("addButton");
add.addEventListener("click", function () {
    if (word.value === "") return;

    showToast("Add Task Complete !");
    addTask();
    saveTasks();
});




const reset = document.getElementById("resetButton");
reset.addEventListener("click", removeText);

function removeText() {
    word.value = "";
    word.focus();
};



function addTask() {
    const newDiv = document.createElement("div");   /*HTML里建立空的<div></div>*/
    newDiv.classList.add("task-item");      /*刚建立空的<div></div>再加上class="task-item"*/



    const newInput = document.createElement("input");   /*HTML里建立空的<input>, 这时还没加入网页的*/
    newInput.type = "checkbox";     /*告诉网页, newInput的类型是checkbox*/
    newDiv.appendChild(newInput);   /*这时变成<div class="task-item"> <input type="checkbox"> </div>*/



    const newSpan = document.createElement("span");     /*HTML里建立空的<span></span>*/
    newSpan.textContent = word.value;                   /*span里面的内容是word.value, 也就是input输入的内容*/
    newDiv.appendChild(newSpan);    /*这时变成<div class="task-item"> <input type="checkbox"> <span>text</span> </div>*/



    newInput.addEventListener("change", function () {   /*当input状态改变的时候, 执行函数*/
        if (newInput.checked) {                         /*当input打钩时*/
            newSpan.classList.add("completed");         /*在span里添加class="completed"*/
        } else {
            newSpan.classList.remove("completed");          /*反之移除class="completed"*/
        }

        saveTasks();                                    /*保存到localStorage*/
    });



    const newButton = document.createElement("button");     /*HTML里建立空的<button></button>*/
    newButton.textContent = "删除";   /*button里添加文字*/
    newButton.classList.add("delete-button");   /*button里class="delete-button"*/
    newDiv.appendChild(newButton);    /*这时变成<div class="task-item"> <input type="checkbox"> <span>text</span> <button class="delete-button">删除</button> </div>*/



    newButton.addEventListener("click", function () {   /*当点击button时*/
        newDiv.remove();                                /*移除当前newDiv*/
        saveTasks();                                    /*保存到localStorage*/
    });


    taskList.appendChild(newDiv);     /*最后, 把刚刚整理好的<div class="task-item"> <input type="checkbox"> <span>text</span> <button class="delete-button">删除</button> </div>, 放进去<div id="task_list">这里</div>里面*/

    word.value = "";    /*清空输入内容*/
    word.focus();       /*这样可以继续输入下一件事, 不用点击*/
};










/*要解释流程*/

const importBtn = document.getElementById("importButton");
const exportBtn = document.getElementById("exportButton");

importBtn.addEventListener("click", importExcel);
exportBtn.addEventListener("click", exportExcel);


/*要解释流程*/
function exportExcel() {

    const tasks = [];

    const items = taskList.querySelectorAll(".task-item");

    items.forEach(function (item) {

        const text = item.querySelector("span").textContent;
        const completed = item.querySelector("input").checked;

        tasks.push({
            Task: text,
            Completed: completed
        });

    });

    const worksheet = XLSX.utils.json_to_sheet(tasks);

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Tasks");

    XLSX.writeFile(workbook, "todo-list.xlsx");
};



/*要解释流程*/
function importExcel() {
    const file = document.getElementById("excelFile").files[0];

    if (!file) {
        return;
    }

    const reader = new FileReader();

    reader.onload = function (event) {

        const data = new Uint8Array(event.target.result);

        const workbook = XLSX.read(data, {
            type: "array"
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const tasks = XLSX.utils.sheet_to_json(sheet);

        console.log(tasks);

        tasks.forEach(function (task) {

            const text = task.Task;
            const completed = task.Completed;

            const newDiv = document.createElement("div");
            newDiv.classList.add("task-item");

            const newInput = document.createElement("input");
            newInput.type = "checkbox";
            newDiv.appendChild(newInput);

            const newSpan = document.createElement("span");
            newSpan.textContent = text;
            newDiv.appendChild(newSpan);

            newInput.checked = completed;
            if (completed) {
                newSpan.classList.add("completed");
            };

            newInput.addEventListener("change", function () {
                if (newInput.checked) {
                    newSpan.classList.add("completed");
                } else {
                    newSpan.classList.remove("completed");
                }
                saveTasks();
            });

            const newButton = document.createElement("button");
            newButton.textContent = "删除";
            newButton.classList.add("delete-button");
            newDiv.appendChild(newButton);

            newButton.addEventListener("click", function () {
                newDiv.remove();
                saveTasks();
            });

            taskList.appendChild(newDiv);
        });
    };

    reader.readAsArrayBuffer(file);
};










function saveTasks() {

    const arr = [];                                             /*创建一个空array*/

    const taskList = document.getElementById("task_list");      /*HTML里找到<div id="task_list"></div>这个容器*/
    const items = taskList.querySelectorAll(".task-item");      /*选择<div id="task_list"></div>里面有class="task-item"的*/

    items.forEach(function (item) {                             /*依序读取items里的每一个task-item, 并暂时存在item这个变量名*/

        const text = item.querySelector("span").textContent;    /*text = 在item里找到span里面的文字*/

        const completed = item.querySelector("input").checked;  /*completed = 在item里找到input是否打钩*/

        arr.push({                                              /*按照以下格式, 把资料存进上面的const arr里*/
            text: text,                                         /*text = 上面的const = text*/
            completed: completed                                /*completed = 上面的const completed*/
        });

    });

    const data = JSON.stringify(arr);                           /*data = 把arr转换成字串string*/

    localStorage.setItem("tasks", data);                        /*在localStorage里储存进名字为tasks, 然后把刚刚转换好的data存进去*/
}





/*要解释流程*/
function loadTasks() {
    const data = localStorage.getItem("tasks");             /*从localStorage里找名字为tasks的资料 (目前还是字串)*/
    if (!data) return;
    const tasks = JSON.parse(data);                         /*把刚刚得到的字串转回Array*/

    tasks.forEach(function (task) {                         /*把tasks里面的每一笔资料拿出来, 暂时命名为task*/

        const text = task.text;                             /**/
        const completed = task.completed;                   /**/

        const newDiv = document.createElement("div");       /**/
        newDiv.classList.add("task-item");                  /**/

        const newInput = document.createElement("input");   /**/
        newInput.type = "checkbox";                         /**/
        newDiv.appendChild(newInput);                       /**/

        const newSpan = document.createElement("span");     /**/
        newSpan.textContent = text;                         /**/
        newDiv.appendChild(newSpan);                        /**/

        newInput.checked = completed;                       /**/
        if (completed) {                                    /**/
            newSpan.classList.add("completed");             /**/
        };

        newInput.addEventListener("change", function () {   /**/
            if (newInput.checked) {                         /**/
                newSpan.classList.add("completed");         /**/
            } else {                                        /**/
                newSpan.classList.remove("completed");      /**/
            }
            saveTasks();                                    /**/
        });


        const newButton = document.createElement("button"); /**/
        newButton.textContent = "删除";                     /**/
        newButton.classList.add("delete-button");           /**/
        newDiv.appendChild(newButton);                      /**/



        newButton.addEventListener("click", function () {   /**/
            newDiv.remove();                                /**/
            saveTasks();                                    /**/
        });

        taskList.appendChild(newDiv);                       /**/
    });

};
loadTasks();                                                /**/


/*要解释流程*/
const toast = document.getElementById("toast");

function showToast(message) {
    toast.textContent = message;
    toast.style.opacity = "1";

    setTimeout(function () {
        toast.style.opacity = "0";
    }, 2000);
}





/*以后可以优化
写一个function专门创建task的,然后之后其它function要用时, 只需写createTask();运行就行*/