// import "./assets/main.css";

import { createApp } from "vue";
import App from "./App.vue";
import vuetify from "./plugins/vuetify";


Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    createApp(App).use(vuetify).mount("#app");
  }
});
