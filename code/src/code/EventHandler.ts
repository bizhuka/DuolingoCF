/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, Excel, navigator, fetch, responsiveVoice */

import * as Utils from "./Utils";

export class EventHandler {
  private _context: Excel.RequestContext | null = null;
  private _sheet: Excel.Worksheet | null = null;
  private imgUrlText: string = "";
  private _audio: HTMLAudioElement | null = null;

  async setHandlers(context: Excel.RequestContext) {
    this._context = context;
    console.log("Setting handlers");

    const worksheets = context.workbook.worksheets.load("items");
    await context.sync();

    for (let i = 0; i < worksheets.items.length; i++)
      worksheets.items[i].tables.load("count");
    await context.sync();

    for (let i = 0; i < worksheets.items.length; i++) {
      const worksheet = worksheets.items[i];

      if (worksheet.tables.count > 0)
        worksheet.onSelectionChanged.add(this.onSelectionChange.bind(this));
    }
    await context.sync();
  }

  async onSelectionChange() {
    if (!this._context) return;
    this._sheet = this._context.workbook.worksheets.getActiveWorksheet();

    const range = this._context.workbook.getActiveCell();

    // Load the address of the selected range
    range.load(["rowIndex", "columnIndex", "values"]);
    await this._context.sync();

    const text = range.values[0][0];
    switch (range.columnIndex) {
      case Utils.IMAGE_COLUMN:
        if (Utils.isValidURL(text)) {
          this._sheet.getRange(
            `${Utils.columnIndexToLetter(Utils.IMAGE_COLUMN)}1`
          ).formulas = [[`=IMAGE("${text}")`]]; // Set the formula to display the image
          await this._context.sync(); // Don't forget to sync after setting the formula
        }
        break;
      case Utils.SOUND_COLUMN:
        if (text) {
          const specificCell = this._sheet.getCell(
            range.rowIndex,
            Utils.SOUND_COLUMN
          );
          specificCell.load("values");
          await this._context.sync();

          const url = specificCell.values[0][0];
          if (!url || !Utils.isValidURL(url)) {
            return;
          }

          if (this._audio) this._audio.pause();
          this._audio = new Audio(url);
          this._audio.play();
        }
        break;
    }
  }

  async onTextBufferImport() {
    if (!this._context) return;

    const newItems: Utils.Card[] = [];

    const textBuffer = await navigator.clipboard.readText();
    const textBufferArr = textBuffer.split("\r\n");

    let newItem: Utils.Card | null = null;
    for (let i = 0; i < textBufferArr.length; i++) {
      const line = textBufferArr[i];
      switch (i % 3) {
        case 0: {
          newItem = Utils.createCard(line);
          newItems.push(newItem);
          break;
        }
        case 1:
          if (newItem) {
            newItem.Back = line;
          }
          break;
        case 2:
          if (line) {
            new Error(`Unexpected line: ${line}`);
          }
          break;
      }
    }

    // Add & push back the new items
    const all_data = await Utils.get_table_data(this._context);
    all_data.push(...newItems);
    Utils.set_table_data(this._context, all_data);
  }

  async onLoadImageFront(root: Utils.IInfo) {
    await this.doLoadImages(root, true);
  }

  async onLoadImageBack(root: Utils.IInfo) {
    await this.doLoadImages(root, false);
  }

  async onLoadSoundFront() {
    await this.doLoadSounds(true);
  }

  async onLoadSoundBack() {
    await this.doLoadSounds(false);
  }

  private async doLoadImages(root: Utils.IInfo, isFront: boolean) {
    if (!this._context) return;

    const all_data = await Utils.get_table_data(this._context);
    for (const card of all_data) {
      if (card.Image) {
        continue;
      }
      const text = isFront ? card.Front : card.Back;
      root.info_text = `Loading image for: ${text}`; // Update the info text

      const allImages = await this.getImageByText(this._context, text);
      // 1st image
      if (allImages && allImages.length > 0) {
        card.Image = allImages[0];
      }
    }
    root.info_text = ""; // Clear the info text

    Utils.set_table_data(this._context, all_data);
  }

  public async getImageByText(context: Excel.RequestContext, text: string) {
    const textToSearch = Utils.removeHTMLTags(text);
    if (!this.imgUrlText) {
      const sheetOption = context.workbook.worksheets.getItem("Opt");
      const imgUrl = sheetOption.getRange("LANG_IMG_URL");
      imgUrl.load("values");
      await context.sync();
      this.imgUrlText = imgUrl.values[0][0];
    }
    console.log(this.imgUrlText + encodeURIComponent(textToSearch));

    const response = await fetch(
      this.imgUrlText + encodeURIComponent(textToSearch)
    );
    if (!response.ok || response.status !== 200) {
      return [""];
    }

    switch (Utils.getDomain(this.imgUrlText)) {
      case "pixabay.com": {
        const data: Utils.PixabayResponse = await response.json();
        if (data.hits.length === 0) {
          return [""];
        }
        return data.hits.map((hit) => hit.previewURL); //webformatURL
      }
      case "www.googleapis.com": {
        const data: Utils.GoogleResponse = await response.json();
        if (data.items.length === 0) {
          return [""];
        }
        return data.items.map((item) =>
          Utils.containsImageExtension(item.link)
            ? item.link
            : item.image.thumbnailLink
        );
      }
    }
  }

  async doLoadSounds(isFront: boolean) {
    if (!this._context) return;

    const sheetOption = this._context.workbook.worksheets.getItem("Opt");
    const soundUrl = sheetOption.getRange("LANG_SOUND_URL");
    soundUrl.load("values");
    await this._context.sync();
    const soundUrlText = soundUrl.values[0][0];

    const all_data = await Utils.get_table_data(this._context);
    for (const card of all_data) {
      if (card.Sound) {
        continue;
      }
      const textToSearch = Utils.removeHTMLTags(
        isFront ? card.Front : card.Back
      );
      card.Sound = soundUrlText + encodeURIComponent(textToSearch);
    }

    Utils.set_table_data(this._context, all_data);
  }

  async onCSVExport() {
    await Excel.run(async (context) => {
      const all_data = await Utils.get_table_data(context);
      const filteredItems = all_data.filter((item) => !item.Exported);

      // Remove the Exported column
      const csvRows = filteredItems.map((item) => {
        const values = Object.values(item);
        values.pop();
        return values;
      });

      Utils.downloadBlob(csvRows, "to_anki.csv");

      all_data.forEach((obj) => (obj.Exported = true));
      Utils.set_table_data(context, all_data);
    });
  }
}
