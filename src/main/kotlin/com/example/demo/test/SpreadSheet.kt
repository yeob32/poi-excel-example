package com.example.demo.test

data class SpreadSheet(
    var name: String,
    var header: List<String>,
    var rows: List<List<String>>
) {

}

data class SpreadSheetMap(
    var name: String,
    var header: List<String>,
    var rows: List<Map<String, String>>
) {

}

class SpreadSheetTest(
    var name: String,
    var header: List<String>,
    var rows: List<SpreadRow>
) {
    class SpreadRow(
        var key: String,
        var value: String
    )
}

fun main() {

}