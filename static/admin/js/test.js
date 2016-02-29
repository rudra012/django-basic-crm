$(function(){
    // $("#tcmodel_set-group").prependTo(".affix");
    $("#tcmodel_set-group table th")[1].style.display = 'none'
    $(".field-followup_date")[0].style.display = 'none'
    temp = $("#tcmodel_set-0-followup_date").val();
    $("#tcmodel_set-0-followup_date").val('');
    function isExist(arr,str) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === str) {
                 return true;
            }
        }
        return false;
    }

    var temp = '';
    var lst = new Array("CB", "PTP");
    $("#id_tcmodel_set-0-calling_code").change(function() {
        if(!isExist(lst, $("#id_tcmodel_set-0-calling_code").val())) {
            $("#tcmodel_set-group table th")[1].style.display = 'none'
            $(".field-followup_date")[0].style.display = 'none'
            temp = $("#tcmodel_set-0-followup_date").val();
            $("#tcmodel_set-0-followup_date").val('');
        }
        else
        {
            $("#tcmodel_set-0-followup_date").val(temp);
            $("#tcmodel_set-group table th")[1].style.display = 'block'
            $(".field-followup_date")[0].style.display = 'block'
        }
    });
});
