#!/usr/bin/env node
"use strict";
const XLSX = require("xlsx");
const { program } = require('commander');

program
  .version('0.0.1')
  .option('-hi, --HelloWorld', 'Hello World').action(function () {


    var arg = process.argv;
    console.log(arg);
    const workbook = XLSX.readFile("./cj.xlsx");
    // 这边直接获取cj文件 没有设置动态获取文件


    var arg = process.argv;
    console.log(arg);
    //获取第一个表单数据
    const first_sheet_name = workbook.SheetNames[0];

    const worksheet1 = workbook.Sheets[first_sheet_name];
    const sheet1 = XLSX.utils.sheet_to_json(worksheet1);
    //获取第二张表单数据
    const second_sheet_name = workbook.SheetNames[1];
    const worksheet2 = workbook.Sheets[second_sheet_name];
    let sheet2 = XLSX.utils.sheet_to_json(worksheet2);

    let reg = /[0-9]{12,}[A-Z]{2}[0-9]{5}\s[\u4e00-\u9fa5]{2,}/g;
    const sheet2New = sheet2
      // .filter((it) => it.DJH == "321081107212JC00071")
      .map((item, i, sheet2) => {
        const itemCopy = Object.assign({}, item);
        ["ZDSZB", "ZDSZD", "ZDSZN", "ZDSZX"].forEach((key) => {
          //字符串 
          let str = item[key];
          // 列出数组的每个元素
          let group = str.match(reg);
          if (group != null) {
            group.forEach((match) => {
              const findItem = sheet1.find((it) => it["原"] == match);

              if (findItem != undefined) {
                // console.log(
                //   match,
                //   findItem["现"],
                //   itemCopy[key].replace(match, findItem["现"])
                // );
                itemCopy[key] = itemCopy[key].replace(match, findItem["现"]);
              }
            });
          }
        });
        // let zdszbArr = item.ZDSZB.match(reg)
        // let zdszdArr = item.ZDSZD.match(reg)
        // let zdsznArr = item.ZDSZN.match(reg)
        // let zdszxArr = item.ZDSZX.match(reg)

        // if (c != null) {
        //     console.log(c);
        // }x
        // console.log("", item, itemCopy);
        return itemCopy;
      });
    // const sheetNew = XLSX.utils.json_to_sheet(sheet2New);
    workbook.Sheets["Sheet2"] = XLSX.utils.json_to_sheet(sheet2New);
    // var ws_name = "SheetJS";
    // XLSX.utils.book_append_sheet(workbook,sheetNew,ws_name);
    XLSX.writeFile(workbook, 'out.xlsx');
    // sheet1.map((item1, j) => {
    //     console.log(j);
    // })

    // item.ZDSZD.match(a)
    // item.ZDSZN.match(a)
    // item.ZDSZX.match(a)

    // for (i = 2; i <= 638; i++) {
    //     //获取E列和F列内容

    //     //获取E列的所有id（Ei）
    //     var address_of_cell = 'E' + i;
    //     //获取Ei的单元格
    //     var desired_cell = worksheet[address_of_cell];
    //     //获取每个Ei单元格的value
    //     var desired_value = (desired_cell ? desired_cell.v : undefined);

    //     var address_of_cell = 'F' + i;
    //     var desired_cell2 = worksheet[address_of_cell];
    //     var desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
    //     //判断E列和F列是否相同，不同就进行替换
    //     if (desired_value != desired_value2) {
    //         let desired_value = desired_value2
    //         desired_value2 = nullx
    //         XLSX.writeFile(workbook, 'out.xlsx');
    //     } else {
    //         console.log('true');
    //     }
    // }



















  })
program.parse(process.argv);
if (program.HelloWorld) console.log("你好世界")
