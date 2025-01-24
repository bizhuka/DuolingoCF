<script setup lang="ts">
declare module '*.vue' {

  import { DefineComponent } from 'vue';

  const component: DefineComponent<{}, {}, any>;

  export default component;

}
import BaseButton from "./components/BaseButton.vue"
</script>

<template>
  <div id="app">
    <div class="content">
      <div class="content-main">
        <h3>Actions</h3>

        <div v-if="!block_main" class="padding">
          <BaseButton :clickHandler="onInitCurrentWorkbook" ButtonText="Init current workbook" />
        </div>

        <div v-if="block_main" class="padding">
          <p v-if="info_text">
            {{ info_text }}
          </p>
          <v-container><v-row justify="center">
              <v-col cols="12">
                <BaseButton :clickHandler="onTextBufferImport" ButtonText="Import from Duolingo buffer" />
                <BaseButton :clickHandler="onLoadImageFront" ButtonText="Load images (Front)" />
                <BaseButton :clickHandler="onLoadImageBack" ButtonText="Load images (Back)" />
                <BaseButton :clickHandler="onFindImageByText" ButtonText="Find image by selected cell" />
                <BaseButton :clickHandler="onLoadSoundFront" ButtonText="Load sounds (Front)" />
                <BaseButton :clickHandler="onLoadSoundBack" ButtonText="Load sounds (Back)" />
                <BaseButton :clickHandler="onCSVExport" ButtonText="Export to CSV" />
              </v-col>
            </v-row>
          </v-container>
        </div>
        <div v-if="images.length">
          <h3>Images</h3>
          <div v-for="(image, index) in images" :key="index">
            <img class="responsive-img" :src="image" @click="handleImageClick($event)">
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script lang="ts">
import { EventHandler } from './code/EventHandler';
import * as Utils from './code/Utils';

export default {
  name: 'Root',
  data: function () {
    return {
      eventHandler: new EventHandler(),
      images: [] as string[],
      block_main: false,
      info_text: '',
    };
  },

  async mounted() {
    this.images = [];

    const _this = this;

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Duo").load("name");
      await context.sync();
      _this.block_main = sheet ? true : false;

      this.eventHandler.setHandlers(context);
    })
  },

  methods: {
    onTextBufferImport() { this.eventHandler.onTextBufferImport(); },
    onLoadImageFront() { this.eventHandler.onLoadImageFront(this); },
    onLoadImageBack() { this.eventHandler.onLoadImageBack(this); },
    onLoadSoundFront() { this.eventHandler.onLoadSoundFront(); },
    onLoadSoundBack() { this.eventHandler.onLoadSoundBack(); },
    onCSVExport() { this.eventHandler.onCSVExport(); },

    async onFindImageByText() {
      this.images = [];

      await Excel.run(async (context) => {
        const activeCell = context.workbook.getActiveCell();
        activeCell.load("values");
        await context.sync();
        if (!activeCell.values) {
          return;
        }
        const allImages = await this.eventHandler.getImageByText(context, activeCell.values[0][0]);

        if (allImages) {
          this.images.push(...allImages);
        }
      });
    },

    async handleImageClick(event: MouseEvent) {
      const src = (event.target as HTMLImageElement)?.src;
      if (!src) {
        return;
      }

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const currentCell = context.workbook.getActiveCell()
        currentCell.load("rowIndex");
        await context.sync();

        sheet.getCell(currentCell.rowIndex, Utils.IMAGE_COLUMN).values = [[src]];
        await context.sync();

        this.images = [];
      });
    },

    async onInitCurrentWorkbook() {
      const _this = this;
      await Excel.run(async (context) => {
        const sheetDuo = context.workbook.worksheets.add("Duo");
        const sheetBack = context.workbook.worksheets.add("Back");
        const sheetOpt = context.workbook.worksheets.add("Opt");

        const headers = [["Front", "Back", "Image", "Hint", "Context", "Sound", "Exported"]];
        this.createTable(sheetDuo, "A2:G2", headers, "DUO_TABLE", true);
        this.createTable(sheetBack, "A2:G2", headers, "BACK_TABLE", true);

        sheetOpt.getRange("A1:B6").values = [
          ["Option", "Value"],
          ["Language", ''],
          ["Domain", ''],
          ["Image search url", ''],
          ["Sound URL", ''],
          ["Voice", '']];
        this.createTable(sheetOpt, "A9:D9", [["Language", "Prefix", "Domain", "https://responsivevoice.org/"]], "LANG_TABLE", false,
          [
            ["Polish", "pl", ".pl", "Polish Male"],
            ["Spanish", "es", ".es", "Spanish Latin American Male"],
            ["Turkish", "tr", ".com.tr", "Turkish Male"],
            ["Russian", "ru", ".ru", "Russian Female"],
            ["English", "en", ".com", "UK English Male"]
          ]
        );
        sheetOpt.names.add("LANG_LIST", "=Opt!A10:A14");

        sheetOpt.names.add("LANG_LANG", "=Opt!B2").getRange().dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "=LANG_LIST"
          }
        }
        sheetOpt.names.add("LANG_DOMAIN", "=Opt!B3").getRange().formulas = [["=VLOOKUP(LANG_LANG,LANG_TABLE,3,FALSE)"]];
        sheetOpt.names.add("LANG_IMG_URL", "=Opt!B4").getRange().formulas = [['="https://pixabay.com/api/?key=44848068-6de875ebe70f9e4ff0373d977&lang=pl&q="']];
        sheetOpt.names.add("LANG_SOUND_URL", "=Opt!B5").getRange().formulas = [['="https://www.google.com/speech-api/v2/synthesize?enc=mpeg&client=chromium&key=AIzaSyBOti4mM-6x9WDnZIjIeyEU21OpBXqWBgw&lang=" & VLOOKUP(LANG_LANG,LANG_TABLE,2,FALSE) & "&text="']];
        sheetOpt.names.add("LANG_VOICE", "=Opt!B6").getRange().formulas = [['=VLOOKUP(LANG_LANG,LANG_TABLE,4,FALSE)']];

        // Delete empty sheets
        const sheets = context.workbook.worksheets;
        sheets.load("items/name"); // Load the names of all sheets to identify them later for deletion
        await context.sync();

        _this.block_main = true;
        _this.eventHandler.setHandlers(context);

        for (let sheet of sheets.items) {
          const usedRange = sheet.getUsedRangeOrNullObject();
          usedRange.load(["values", "isNullObject"]);
          await context.sync();

          if (usedRange.isNullObject && !usedRange.values) {
            sheet.delete();
          }
        }
      });
    },

    async createTable(sheet: Excel.Worksheet, where: string, headers: any[][], tableName: string, freeze: boolean, values: string[][] = []) {
      // Get a range starting from the second row and spanning the number of columns in headers
      const headerRange = sheet.getRange(where);
      headerRange.values = headers;

      // Create a table with the headers
      const table = sheet.tables.add(where, true /*hasHeaders*/);
      table.name = tableName;
      table.getHeaderRowRange().values = headers;

      if (freeze) {
        sheet.freezePanes.freezeRows(2);
      }

      if (values && values.length) {
        table.rows.add(-1, values);
      }
      sheet.getCell(0, 0).select();
    },
  }
}
</script>



<style></style>