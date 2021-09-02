namespace QunitTests {

    var QUnit = Library.QUnitGS2.QUnit;

    function divideThenRound(numberator, denominator) {
        return Math.round(numberator / denominator);
    }

    export function tesfsForQunit() {
        QUnit.module("Basic tests")

        QUnit.test("simple numbers", function( assert ) {
            assert.equal(divideThenRound(10, 2), 5, "whole numbers");
            assert.equal(divideThenRound(10, 4), 3, "decimal numbers");
        });

        QUnit.start();
    }
}