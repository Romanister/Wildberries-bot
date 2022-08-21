const { spawn } = require("child_process");
const {
  mouse,
  left,
  right,
  up,
  down,
  Point,
  keyboard,
  Key,
  clipboard,
} = require("@nut-tree/nut-js");
const util = require("util");
const puppeteer = require("puppeteer");
const path = require("path");
const formData = require("form-data");
const form = new formData();
const fs = require("fs").promises;
let xlsx = require("xlsx");
const { access } = require("fs");
const { default: axios } = require("axios");
const FILENAME = "123.xlsx";

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
function ValidData(massiv) {
  let c = [];
  for (const value of massiv) {
    if (value.hasOwnProperty("Артикул товара")) {
      c.push(value);
    }
  }
  return c;
}
(async function () {
  let file = xlsx.readFile(FILENAME);
  let sheetNames = file.SheetNames;
  let data = xlsx.utils.sheet_to_json(file.Sheets[sheetNames[0]]);
  let validArtic = ValidData(data);
  for (let i = 0; i < 1; i++) {
    const browser = await puppeteer.launch({
      headless: false,
      ignoreHTTPSErrors: true,
      args: [`--window-size=1920,1080`],
      defaultViewport: {
        width: 1920,
        height: 1080,
      },
    });
    const page = await browser.newPage();
    const cookieJson = await fs.readFile("cookies.json");
    const cookies = JSON.parse(cookieJson);
    await page.setCookie(...cookies);
    await page.goto(
      `https://seller.wildberries.ru/new-goods/product-card/card?vendorCode=${validArtic[i]["Артикул товара"]}`
    );
    await mouse.setPosition(new Point(1920 / 2, 1080 / 2));
    await mouse.leftClick();
    await sleep(10000);
    await keyboard.pressKey(Key.LeftControl, Key.F);
    await keyboard.releaseKey(Key.LeftControl, Key.F);
    await clipboard.copy("MOV");
    await keyboard.pressKey(Key.LeftControl, Key.V);
    await keyboard.releaseKey(Key.LeftControl, Key.V);
    await sleep(500);
    await mouse.setPosition(new Point(810, 585));
    await mouse.leftClick();
    await sleep(500);
    await mouse.setPosition(new Point(445, 60));
    await mouse.leftClick();
    await clipboard.copy("C:\\Users\\roman\\Desktop\\Foto");
    await keyboard.pressKey(Key.LeftControl, Key.V);
    await keyboard.releaseKey(Key.LeftControl, Key.V);
    await keyboard.pressKey(Key.Enter);
    await keyboard.releaseKey(Key.Enter);
    await sleep(1500);
    await mouse.setPosition(new Point(956, 60));
    await mouse.leftClick();
    await sleep(500);
    require("child_process")
      .spawn("clip")
      .stdin.end(util.inspect(validArtic[i]["Артикул товара"]));
    await keyboard.pressKey(Key.LeftControl, Key.V);
    await keyboard.releaseKey(Key.LeftControl, Key.V);
    await keyboard.pressKey(Key.Enter);
    await keyboard.releaseKey(Key.Enter);
    await sleep(1500);
    await mouse.setPosition(new Point(750, 300));
    await mouse.leftClick();
    await keyboard.pressKey(Key.LeftControl, Key.A);
    await keyboard.releaseKey(Key.LeftControl, Key.A);
    await sleep(500);
    await keyboard.pressKey(Key.Enter);
    await keyboard.releaseKey(Key.Enter);
    await sleep(7000);
    await keyboard.pressKey(Key.LeftControl, Key.F);
    await keyboard.releaseKey(Key.LeftControl, Key.F);
    await clipboard.copy("wb");
    await keyboard.pressKey(Key.LeftControl, Key.V);
    await keyboard.releaseKey(Key.LeftControl, Key.V);
    await sleep(500);
    await mouse.setPosition(new Point(145, 500));
    await mouse.leftClick();
    await sleep(1000);
    await browser.close();
  }
})();
