if (document.querySelector(".path")) {
    var path = document.querySelector(".path");
    path.querySelector("li.displaynone").previousElementSibling.classList.add("current");
}

function removePagingArea(oTarget) {
    if ($(oTarget).length < 1 && (oTarget != "#prdReview" || oTarget != "#prdQna")) return;

    if ($(oTarget).css("display") == "block") {
        if (oTarget == "#prdReview") {
            var record = $(".xans-record-:first", ".xans-product-review");
            if (record.length < 1 || record.is(":not(:visible)")) {
                $(".xans-product-reviewpaging").remove();
            }
        } else if (oTarget == "#prdQnA") {
            var record = $(".xans-record-:first", ".xans-product-qna");
            if (record.length < 1 || record.is(":not(:visible)")) {
                $(".xans-product-qnapaging").remove();
            }
        }
    }
}

$(document).ready(function () {
    $("#actionCartClone, #actionWishClone, #actionBuyClone, #actionWishSoldoutClone")
        .unbind()
        .bind("click", function () {
            try {
                var id = $(this).attr("id").replace(/Clone/g, "");
                if (typeof id !== "undefined") $("#" + id).trigger("click");
                else return false;
            } catch (e) {
                return false;
            }
        });

    function productDetailOrigin() {
        var imgChk = $("#prdDetailContent").find("img").length;
        var thumbSrc = $(".jsThumbnail").data("src");
        $(".thumbnail__item").eq(0).find("img").attr("src", thumbSrc);
        $(".xans-product-detail .imgArea .listImg li").eq(0).find("img").attr("src", thumbSrc);
        if (imgChk <= 0) {
            $("#prdDetailBtn").remove();
        }
    }
    productDetailOrigin();

    // Add Image
    var oTarget = $(".xans-product-mobileimage ul li");
    var oAppend = oTarget.first().children("p").clone();

    oTarget.not(":first").each(function () {
        $(this)
            .children()
            .wrap(function () {
                return '<p class="thumbnail">' + $(this).html() + oAppend.html() + "</p>";
            });

        $(this).children("p").children("img:first").remove();
    });
});
