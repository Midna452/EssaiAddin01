Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    Word.run(async (context) => {
      const cleared = await clearText(context);
      await context.sync();

      if(cleared){
        // Insertion automatique au démarrage
        await insertInitialMessage(["Cybersécurité", "Réseau", "Développement", "Les maths dans le développement", "Conclusion"]);
      }else{}
    });
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    correction();
    await context.sync();
  });
}

async function clearText(context) {
  const body = context.document.body;
  body.load("text");
  await context.sync();

  if (body.text !== "") {
    return new Promise((resolve) => {
      document.getElementById("modal").style.display = "flex";

      document.getElementById("btnYes").onclick = async () => {
        await Word.run(async (ctx) => {
          ctx.document.body.clear();
          await ctx.sync();
        });
        document.getElementById("modal").style.display = "none";
        resolve(true);
      };

      document.getElementById("btnNo").onclick = () => {
        document.getElementById("modal").style.display = "none";
        resolve(false);
      };
    });
  } else {
    return Promise.resolve(true); // Rien à supprimer, on continue
  }
}

async function insertInitialMessage(listeTitre) {
  await Word.run(async (context) => {
    const body = context.document.body;

    for (let i = 1; i < 5; i++) {
      // Titre
      const title = body.insertParagraph(`${listeTitre[i]}`, Word.InsertLocation.end);
      title.style = "Titre 1";
      await context.sync();

      // Sous-titre 1
      const subtitle1 = body.insertParagraph(`Section - Présentation générale`, Word.InsertLocation.end);
      subtitle1.style = "Titre 2";
      await context.sync();

      // Paragraphe texte normal
      const para1 = body.insertParagraph(
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus luctus urna sed urna ultricies ac tempor dui sagittis. Suspendisse auctor, leo ac consequat faucibus, ipsum tortor fermentum sem, nec cursus tellus urna et elit. Curabitur in orci eget nunc dictum hendrerit non vitae enim. Maecenas mattis tristique quam id sollicitudin. Aenean consectetur arcu et nunc bibendum, at maximus sem condimentum. Mauris ac accumsan magna, a hendrerit leo. Maecenas ullamcorper urna nibh, vel sollicitudin tellus pellentesque in. Mauris tempus scelerisque mi, vitae imperdiet arcu. In auctor condimentum nunc vitae scelerisque. Aenean leo nulla, blandit quis finibus volutpat, auctor volutpat purus. Suspendisse tincidunt ipsum id ligula lacinia pharetra. Proin bibendum tortor sit amet sem pretium, quis ultrices libero volutpat. Maecenas eget luctus elit",
        Word.InsertLocation.end
      );
      para1.style = "Normal"; // Forcer style normal ou laisse cette ligne, normalement déjà normal

      // Sous-titre 2
      const subtitle2 = body.insertParagraph(`Section - Détails spécifiques`, Word.InsertLocation.end);
      subtitle2.style = "Titre 2";
      await context.sync();

      // Paragraphe texte normal
      const para2 = body.insertParagraph(
        "In condimentum facilisis porta. Sed nec diam eu diam mattis viverra. Nulla fringilla, orci ac euismod semper. Suspendisse auctor, leo ac consequat faucibus, ipsum tortor fermentum sem, nec cursus tellus urna et elit. Curabitur in orci eget nunc dictum hendrerit non vitae enim. Maecenas mattis tristique quam id sollicitudin. Aenean consectetur arcu et nunc bibendum, at maximus sem condimentum. Mauris ac accumsan magna, a hendrerit leo. Maecenas ullamcorper urna nibh, vel sollicitudin tellus pellentesque in. Mauris tempus scelerisque mi, vitae imperdiet arcu. In auctor condimentum nunc vitae scelerisque. Aenean leo nulla, blandit quis finibus volutpat, auctor volutpat purus. Suspendisse tincidunt ipsum id ligula lacinia pharetra. Proin bibendum tortor sit amet sem pretium, quis ultrices libero volutpat. Maecenas eget luctus elit",
        Word.InsertLocation.end
      );
      para2.style = "Normal"; // Forcer style normal

      // Saut de page sauf dernière page
      if (i < 4) {
        body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
      }

      await context.sync();
    }
  });
}

async function correction() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style");
    await context.sync();

    const stylesAcceptes = ["Titre 1", "Titre 2", "Titre 3"];
    const stylesDetectes = new Set();

    paragraphs.items.forEach(paragraph => {
      if (stylesAcceptes.includes(paragraph.style)) {
        stylesDetectes.add(paragraph.style);
      }
    });

    const stylesUtilises = Array.from(stylesDetectes);

    const popup = document.getElementById("correction");
    const popupText = document.getElementById("resCorrection");
    popup.style.display = "flex";

    if (stylesUtilises.length === 0) {
      console.log("Aucun titre avec style 'Titre 1', 'Titre 2' ou 'Titre 3' détecté.");
      popupText.innerText = "Résultat : ❌";
    } else {
      console.log("Styles titres détectés :", stylesUtilises);
      popupText.innerText = "Résultat : ✔️";
    }

    const btn = document.getElementById("btnValidation");
    if (btn) {
      btn.onclick = () => {
        popup.style.display = "none";
      };
    } else {
      console.error("Bouton btnValidation introuvable dans le DOM.");
    }
  });
}



