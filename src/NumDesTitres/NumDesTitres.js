/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    await clearText(context)
    await context.sync();
  });
}

async function clearText(context) {
  const body = context.document.body;
  body.load("text"); // Charger le texte du document

  await context.sync();

  if (body.text !== "") {
    document.getElementById("modal").style.display = "flex";

    document.getElementById("btnYes").onclick = async () => {
      body.clear();
      await context.sync();
      document.getElementById("modal").style.display = "none";
      alert("Le document a été effacé !");
    };

    document.getElementById("btnNo").onclick = () => {
      document.getElementById("modal").style.display = "none";
    };
  } else {
    alert("Le document est déjà vide.");
  }
}

