// editar.js
let idx;          // será preenchido pelo HTML
function addItem(){
  const tbody = document.querySelector("#tblItens tbody");
  const tr = document.createElement("tr");

  const material = document.createElement("input");
  material.type = "text";
  material.name = "material_" + idx;
  material.required = true;

  const qtde = document.createElement("input");
  qtde.type = "number";
  qtde.name = "qtde_" + idx;
  qtde.min = 1;
  qtde.required = true;

  const obs = document.createElement("input");
  obs.type = "text";
  obs.name = "obs_item_" + idx;

  const btn = document.createElement("button");
  btn.type = "button";
  btn.textContent = "Remover";
  btn.onclick = function () { removeItem(btn); };

  const tdMaterial = document.createElement("td");
  tdMaterial.appendChild(material);
  const tdQtde = document.createElement("td");
  tdQtde.appendChild(qtde);
  const tdObs = document.createElement("td");
  tdObs.appendChild(obs);
  const tdBtn = document.createElement("td");
  tdBtn.appendChild(btn);

  tr.appendChild(tdMaterial);
  tr.appendChild(tdQtde);
  tr.appendChild(tdObs);
  tr.appendChild(tdBtn);

  tbody.appendChild(tr);
  document.getElementById("total_itens").value = ++idx;
}

function removeItem(btn){
  btn.closest("tr").remove();
  document.getElementById("total_itens").value = document.querySelectorAll("#tblItens tbody tr").length;
}

function enviar(){
  const usuario = document.getElementById("usuario").value.trim();
  if (!usuario){
    alert("Informe quem está editando");
    return false;
  }
  return true;
}