import { join, resolve } from "path";
import dotenv from "dotenv";
import mysql from "mysql2/promise";
import fs from "fs";
import { Table } from "console-table-printer";
import dayjs from "dayjs";
// https://github.com/exceljs/exceljs
import ExcelJS from "exceljs";
import { exit } from "process";
import shelljs from "shelljs";
import minimist from "minimist";

// 현재 작업 디렉토리
const cwd = shelljs.pwd().toString().replaceAll("\\", "/");

// 명령줄 파라미터
const {d, h, u, p, output, port} = minimist(process.argv.slice(2));

// 연동정보 설정
const database = d || "test";
const hostname = h || "127.0.0.1";
const username = u || "root";
const password = p || "123qwe!@#";
const outputDir = output || cwd;
const portNumber = port || 3306;

// console.log("============================");
// console.log(`hostname: ${hostname}`);
// console.log(`port-number: ${portNumber}`);
// console.log(`database: ${database}`);
// console.log(`username: ${username}`);
// console.log(`password: ${password}`);
// console.log(`outputDir: ${outputDir}`);
// console.log("============================");

function message(msg, ms = 500, table = false) {
    if (table) {
        const p = new Table();

        for (let v of msg) {
            p.addRow(v);
        }

        p.printTable();
    } else {
        console.log(msg);
    }

    return new Promise((resolve) => {
        setTimeout(resolve, ms);
    });
}

// 생성될 파일의 이름을 지정한다.
const outputFileName = `${outputDir}/${database}_테이블명세서_${dayjs().format("YYMMDD_HHmmss")}.xlsx`;

// 접속 정보 설정
const connectionInfo = {
    host: hostname, // MYSQL 서버 주소 (다른 PC인 경우 IP주소),
    port: portNumber, // MYSQL 포트번호
    user: username, // MYSQL의 로그인 할 수 있는 계정이름
    password: password, // 비밀번호
    database: database, // 사용하고자 하는 데이터베이스 이름
};

console.log("================================================");
console.log("|          Configuration Information           |");
console.log("================================================");

for (let key in connectionInfo) {
    console.log(`- ${key}: ${connectionInfo[key]}`);
}

const headerStyle = {
    font: { name: "맑은 고딕", size: 11, bold: true },
    fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00FFFF00" },
        bgColor: { argb: "00FFFF00" },
    },
    border: {
        top: { style: "thin", color: { argb: "00000000" } },
        left: { style: "thin", color: { argb: "00000000" } },
        bottom: { style: "thin", color: { argb: "00000000" } },
        right: { style: "thin", color: { argb: "00000000" } },
    },
    alignment: {
        vertical: "middle",
        horizontal: "center",
    },
};

const bodyStyle = {
    font: { name: "맑은 고딕", size: 11 },
    border: {
        top: { style: "thin", color: { argb: "000000" } },
        left: { style: "thin", color: { argb: "000000" } },
        bottom: { style: "thin", color: { argb: "000000" } },
        right: { style: "thin", color: { argb: "000000" } },
    },
    fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00FFFFFF" },
        bgColor: { argb: "00FFFFFF" },
    },
    alignment: {
        vertical: "middle",
        horizontal: "center",
    },
};

(async () => {
    let dbcon = null;
    let sql = null;
    let input = null;
    let tableCount = 0;
    let step = 1;

    try {
        await message(`\nstep${step++}: 데이터베이스에 접속합니다...`);
        dbcon = await mysql.createConnection(connectionInfo);
        await dbcon.connect();
        await message(" >> 데이터베이스 접속 성공");

        await message(`\nstep${step++}: 분석결과가 저장될 파일을 생성합니다.`);

        const workbook = new ExcelJS.Workbook();
        workbook.creator = "이광호강사(leekh4232@gmail.com)";
        workbook.lastModifiedBy = "이광호강사(leekh4232@gmail.com)";
        workbook.created = new Date();
        workbook.modified = new Date();
        workbook.lastPrinted = new Date();
        workbook.calcProperties.fullCalcOnLoad = true;
        workbook.views = [
            {
                x: 0,
                y: 0,
                firstSheet: 0,
                activeTab: 0,
                visibility: "visible",
            },
        ];

        const sheet1 = workbook.addWorksheet("테이블명세서");
        sheet1.getColumn("A").width = 7;
        sheet1.getColumn("B").width = 18;
        sheet1.getColumn("C").width = 14;
        sheet1.getColumn("D").width = 12;
        sheet1.getColumn("E").width = 8;
        sheet1.getColumn("F").width = 16;
        sheet1.getColumn("G").width = 12;
        sheet1.getColumn("H").width = 28;

        await message(` >> ${outputFileName} 파일이 생성됨`);

        await message(`\nstep${step++}: 테이블 목록을 조회합니다.`);
        sql = "SELECT table_name as `name`, TABLE_COMMENT as `comment` FROM information_schema.tables WHERE table_schema=?";
        input = [connectionInfo.database];
        const [tableList] = await dbcon.query(sql, input);
        tableCount = tableList.length;
        await message(` >> ${tableCount}개의 테이블이 검색되었습니다.`);

        if (tableCount < 1) {
            throw new Error("현재 데이터베이스에 접근 가능한 테이블이 없습니다.");
        }

        sql = `SELECT
			ORDINAL_POSITION AS No,
			COLUMN_NAME AS 필드명,
			COLUMN_TYPE AS 데이터타입,
			if( IS_NULLABLE = 'NO', 'NOT NULL', 'NULL' ) as 널허용,
			COLUMN_KEY AS 키,
			EXTRA AS 옵션,
			COLUMN_DEFAULT 기본값,
			COLUMN_COMMENT AS 설명
		FROM
			INFORMATION_SCHEMA.COLUMNS
		WHERE
			TABLE_SCHEMA = ?
			AND TABLE_NAME = ?
		ORDER BY
			TABLE_NAME, ORDINAL_POSITION`;

        let currentRow = 1;

        for (let i = 0; i < tableCount; i++) {
            await message(`\nstep${step++}: ${tableList[i].name} 테이블의 구조를 분석합니다.`);
            await message(`- 테이블 설명: ${tableList[i].comment}`);

            sheet1.mergeCells(`A${currentRow}:B${currentRow}`);
            sheet1.getCell(`A${currentRow}`).value = "테이블 이름";
            sheet1.getCell(`A${currentRow}`).alignment = { horizontal: "center", vertical: "middle" };
            sheet1.getCell(`A${currentRow}`).style = headerStyle;
            sheet1.getCell(`B${currentRow}`).style = headerStyle;
            sheet1.getCell(`A${currentRow}`).fill = headerStyle.fill;
            sheet1.getCell(`B${currentRow}`).fill = headerStyle.fill;

            sheet1.mergeCells(`C${currentRow}:H${currentRow}`);
            sheet1.getCell(`C${currentRow}`).value = tableList[i].name;
            sheet1.getCell(`C${currentRow}`).alignment = { horizontal: "left", vertical: "middle" };
            sheet1.getCell(`C${currentRow}`).alignment = {
                vertical: "middle",
                horizontal: "left",
            };
            sheet1.getCell(`C${currentRow}`).style = bodyStyle;
            sheet1.getCell(`D${currentRow}`).style = bodyStyle;
            sheet1.getCell(`E${currentRow}`).style = bodyStyle;
            sheet1.getCell(`F${currentRow}`).style = bodyStyle;
            sheet1.getCell(`G${currentRow}`).style = bodyStyle;
            sheet1.getCell(`H${currentRow}`).style = bodyStyle;
            sheet1.getCell(`C${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`D${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`E${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`F${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`G${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`H${currentRow}`).fill = bodyStyle.fill;
            currentRow++;

            sheet1.mergeCells(`A${currentRow}:B${currentRow}`);
            sheet1.getCell(`A${currentRow}`).value = "테이블 설명";
            sheet1.getCell(`A${currentRow}`).alignment = { horizontal: "center", vertical: "middle" };
            sheet1.getCell(`A${currentRow}`).style = headerStyle;
            sheet1.getCell(`B${currentRow}`).style = headerStyle;
            sheet1.getCell(`A${currentRow}`).fill = headerStyle.fill;
            sheet1.getCell(`B${currentRow}`).fill = headerStyle.fill;

            sheet1.mergeCells(`C${currentRow}:H${currentRow}`);
            sheet1.getCell(`C${currentRow}`).alignment = {
                vertical: "middle",
                horizontal: "left",
            };
            sheet1.getCell(`C${currentRow}`).value = tableList[i].comment;
            sheet1.getCell(`C${currentRow}`).alignment = { horizontal: "left", vertical: "middle" };
            sheet1.getCell(`C${currentRow}`).style = bodyStyle;
            sheet1.getCell(`C${currentRow}`).style = bodyStyle;
            sheet1.getCell(`C${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`D${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`E${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`F${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`G${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`H${currentRow}`).fill = bodyStyle.fill;
            sheet1.getCell(`H${currentRow}`).style = bodyStyle;
            currentRow++;

            sheet1.addRow(["No", "필드명", "데이터타입", "널허용", "키", "옵션", "기본값", "설명"]);
            for (let k of ["A", "B", "C", "D", "E", "F", "G", "H"]) {
                sheet1.getCell(`${k}${currentRow}`).style = headerStyle;
                sheet1.getCell(`${k}${currentRow}`).fill = headerStyle.fill;
            }
            currentRow++;

            const [result] = await dbcon.query(sql, [connectionInfo.database, tableList[i].name]);
            await message(result, 0, true);
            for (let row of result) {
                const items = [];

                for (let col in row) {
                    items.push(row[col]);
                }

                sheet1.addRow(items);

                for (let k of ["A", "B", "C", "D", "E", "F", "G", "H"]) {
                    sheet1.getCell(`${k}${currentRow}`).style = bodyStyle;
                    sheet1.getCell(`${k}${currentRow}`).fill = bodyStyle.fill;
                }

                sheet1.getCell(`H${currentRow}`).alignment = {
                    vertical: "middle",
                    horizontal: "left",
                };

                currentRow++;
            }

            sheet1.mergeCells(`A${currentRow}:H${currentRow}`);
            sheet1.getCell(`A${currentRow}`).value = "";
            currentRow++;

            sheet1.mergeCells(`A${currentRow}:H${currentRow}`);
            sheet1.getCell(`A${currentRow}`).value = "";
            currentRow++;

            // await message(`\nDROP TABLE IF EXISTS ${tableList[i].name};\n`);
            // const [struct] = await dbcon.query("show create table " + [tableList[i].name]);
            // const query = struct[0]['Create Table'];
            // await message(query + "\n", 1000);
        }

        await message(`\nstep${step++}: 분석결과를 저장합니다.`);

        await workbook.xlsx.writeFile(outputFileName);
        await message(` >> 분석결과가 ${outputFileName}에 저장됨\n`, 1000);
    } catch (err) {
        console.error(err);
        return;
    } finally {
        if (dbcon) {
            dbcon.end();
        }
        await message(`프로그램을 종료합니다. :)`);

        process.exit(1);
    }
})();
