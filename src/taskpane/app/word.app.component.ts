import { Component } from '@angular/core';

/* global Word */

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Office Word Add-ins feat Angular 🚀", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "red";
      paragraph.font.size = 40;
      paragraph.font.bold = true;
      paragraph.alignment=Word.Alignment.centered;
      await context.sync();
    });
  }
}
