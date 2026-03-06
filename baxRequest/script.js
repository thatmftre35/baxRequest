let equipmentData = [];
let equipmentCount = 0;

fetch('equipment.xlsx')
.then(res => res.arrayBuffer())
.then(data => {
    let workbook = XLSX.read(data);
    let sheet = workbook.Sheets[workbook.SheetNames[0]];
    equipmentData = XLSX.utils.sheet_to_json(sheet);
});

document.getElementById("requestType").addEventListener("change",function(){
    document.getElementById("equipmentSection").innerHTML = "";
    equipmentCount = 0;
});

function addEquipment(){

let section = document.getElementById("equipmentSection");
let type = document.getElementById("requestType").value;

if(!type){
alert("Select Owned or Rental first");
return;
}

equipmentCount++;

let div = document.createElement("div");
div.className = "equipmentItem";

if(type === "owned"){

let types = [...new Set(equipmentData.map(e=>e.Type))];

let typeOptions = types.map(t=>`<option>${t}</option>`).join("");

div.innerHTML = `

<label>Equipment Type</label>
<select onchange="loadModels(this,${equipmentCount})" id="type${equipmentCount}">
<option>Select</option>
${typeOptions}
</select>

<label>Model (optional)</label>
<select id="model${equipmentCount}">
<option>Select</option>
</select>

`;

}else{

div.innerHTML = `

<label>Equipment Type</label>
<input id="type${equipmentCount}" placeholder="Equipment Type">

<label>Model (optional)</label>
<input id="model${equipmentCount}" placeholder="Model">

`;

}

section.appendChild(div);

}

function loadModels(select,id){

let type = select.value;

let models = equipmentData
.filter(e=>e.Type===type)
.map(e=>e.Model);

let dropdown = document.getElementById(`model${id}`);

dropdown.innerHTML = "<option>Select</option>";

models.forEach(m=>{
dropdown.innerHTML += `<option>${m}</option>`;
});

}

function submitForm(){

let name = document.getElementById("name").value;
let project = document.getElementById("project").value;
let required = document.getElementById("required").value;
let ret = document.getElementById("return").value;
let notes = document.getElementById("notes").value;

let equipment = [];

for(let i=1;i<=equipmentCount;i++){

let type = document.getElementById(`type${i}`)?.value;
let model = document.getElementById(`model${i}`)?.value;

if(type){
equipment.push(type + " - " + model);
}

}

let body = `
Name: ${name}

Project: ${project}

Required By: ${required}

Return: ${ret}

Equipment:
${equipment.join("\n")}

Notes:
${notes}
`;

window.location.href = `mailto:equipment@company.com?subject=Equipment Request&body=${encodeURIComponent(body)}`;

}