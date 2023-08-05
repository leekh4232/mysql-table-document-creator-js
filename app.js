#!/usr/bin/env node

import shelljs from "shelljs";
import minimist from "minimist";
import createTableDocument from "./createTableDocument.js";

// 현재 작업 디렉토리
const cwd = shelljs.pwd().toString();

// 명령줄 파라미터
const {d, h, u, p, output, port} = minimist(process.argv.slice(2));

// DATABASE 연동정보 설정
const env = {
    host : h || "127.0.0.1",
    port : port || 3306,
    user : u || "root",
    password : p || "123qwe!@#",
    database : d || "myschool",
    output : output || cwd,
    connectionLimit: 10,
    connectTimeout: 30000,
    waitForConnections: true
};

// 프로그램 시작
console.clear();
console.log("================================================");
console.log("|         MySQL DATABASE Util (by leekh)       |");
console.log("================================================");

for (let key in env) {
    console.log(`- ${key}: ${env[key]}`);
}


createTableDocument(env);




