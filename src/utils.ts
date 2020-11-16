﻿namespace Utils{
    /**
     * 2차원 배열에서 일치하는 값을 찾아 Row 인덱스를 반환하는 함수 
     *  
     * @param {any[][]} data - 검색할 2차원 배열
     * @param {string} inColumn - 검색할 컬럼 헤더 
     * @param {any} value - 검색 대상 값 
     * @param {boolean} isSame - 찾는 값이 맞는지 확인하여 boolean값을 반환하는 함수
     * @return {number} - row index를 반환
     */
    export function indexOfinDate(data, inColumn, value, isSame) {
        const r = -1;
        let columnIndex;
        let startRow;
        if (data.length > 0) {
            const inColumnType = typeof inColumn;
            switch (inColumnType) {
                case 'number':
                    columnIndex = inColumn;
                    if (columnIndex > data[0].length) {
                        Library.Logger.severe(`columnIndex(${columnIndex})이 범위 밖입니다.`);
                        return -1;
                    }
                    startRow = 0; // 컬럼인덱스가 숫자로 들어올 경우 제목줄 없는것으로 판단 처음부터검색
                case 'string':
                    columnIndex = data[0].indexOf(inColumn);
                    startRow = 1; // 컬럼인덱스가 문자열로 들어올 경우 제목줄 포함이므로 다음줄부터 검색
                    break;
                default:
                    Library.Logger.severe(`컬럼타입(${inColumnType})이 잘못입력되었습니다.`);
                    return r;
            }
            for (let i = startRow; i < data.length; i++) {
                Library.Logger.log(`${startRow}, ${data.length}, ${i}, ${data[i][columnIndex]}, ${value}`);
                if (data[startRow][0] == undefined) {
                    if (isSame(data[i], value))
                        return i;
                }
                else {
                    if ( //columnIndex < 0 && isSame(data[i], value) ||
                    columnIndex >= 0 && isSame(data[i][columnIndex], value))
                        return i;
                }
            }
            return r;
        }
        else {
            return data;
        }
    }

    /**
     *  2차 배열을 tanspose 해주는 함수
     *  a =
     *  [
     *      ["a", "b", "c"],
     *      ["d", "e", "f"],
     *  ]
     *  a[0] = {0:"a", 1:"b", 2:"c"}
     *  Object.keys(a[0]) = [0, 1, 2]
     *
     *  이 곳에서 복사해 옴 : https://stackoverflow.com/a/16705104/9457247
     * @param  {Object[][]} a 입력받는 2D Array 데이타
     * @returns {Object[][]} transpose된 2D Array 데이타
     */
    function transpose_array(a) {
        return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
    }

    /**
     *  transpose할 range를 입력받아 바로 그자리에서 transpose하고 해당 Range를 반환함.
     *
     * @param {Types.Range} inRange 입력받는 Range 객체
     * @returns {Types.Range} transpose되어 출력되는 Range 객체
     */
    export function transpose_range(inRange) {
        var outRange = inRange.getSheet().getRange(inRange.getRow(), inRange.getColumn(), inRange.getNumColumns(), inRange.getNumRows());
        var outValues = transpose_array(inRange.getValues());
        inRange.clear();
        outRange.setValues(outValues);
        return outRange;
    }

}