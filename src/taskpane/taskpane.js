Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // run the script when a doc is opened
    run();
  }
});

async function run() {
  try {
    await Word.run(async (context) => {
      // Load document's paragraphs
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");

      // Sync to apply the load
      await context.sync();

      // loop through the paragraphs
      paragraphs.items.forEach((paragraph) => {
        if (paragraph.style === "Heading 1") {
          paragraph.font.color = "red"; // Change this to your desired color
        }
      });

      // Sync to apply the changes
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
