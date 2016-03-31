$(function(){
    $("#tcmodel_set-group").insertBefore(".submit-row");
    $(".form-expand").hide();
    $("#tcmodel_set-group")[0].style.margin = '6px 0px';
//    $("#tcmodel_set-group table")[0].style.width = '96%';
    $("#tcmodel_set-group table")[0].style.marginBottom = '10px';

    $("#tcmodel_set-group table th")[0].style.display = 'block';
    $("#tcmodel_set-group table th")[2].style.display = 'block';

    $("#tcmodel_set-group table th")[2].style.position = 'absolute';
    $("#tcmodel_set-group table th")[2].style.marginTop = '77px';
    $("#tcmodel_set-group table th")[2].style.width = '100%';
    $("#tcmodel_set-group table td")[3].style.marginTop = '50px';

    $("#tcmodel_set-group table td")[0].style.display = 'block';
    $("#tcmodel_set-group table td")[1].style.display = 'block';
    $("#tcmodel_set-group table td")[3].style.display = 'block';


    $("#tcmodel_set-group table th")[1].style.display = 'none';
    $("#tcmodel_set-0 .field-followup_date")[0].style.display = 'none';
    $("#tcmodel_set-2-group .original p").hide();
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
    var lst = new Array("CB", "PTP", "WPD", "OS");

    function hideshowCBDate()
    {
        if(!isExist(lst, $("#id_tcmodel_set-0-calling_code").val())) {
                $("#tcmodel_set-group table th")[1].style.display = 'none';
                $("#tcmodel_set-0 .field-followup_date")[0].style.display = 'none';
                temp = $("#tcmodel_set-0-followup_date").val();
                $("#tcmodel_set-0-followup_date").val('');

                $("#tcmodel_set-group table th")[2].style.marginTop = '77px';
        }
        else
        {
            $("#tcmodel_set-0-followup_date").val(temp);
            $("#tcmodel_set-group table th")[1].style.display = 'block';
            $("#tcmodel_set-0 .field-followup_date")[0].style.display = 'block';

            $("#tcmodel_set-group table th")[2].style.marginTop = '177px';

            $("#tcmodel_set-group table th")[1].style.position = 'absolute';
            $("#tcmodel_set-group table th")[1].style.marginTop = '77px';
            $("#tcmodel_set-group table th")[1].style.width = '100%';
            $("#tcmodel_set-group table td")[2].style.marginTop = '50px';
        }
    }
    hideshowCBDate();

    $("#id_tcmodel_set-0-calling_code").change(function() {
        hideshowCBDate();
    });

    $( "#suprememodel_form" ).submit(function() {
        flag = true;
        if(isExist(lst, $("#id_tcmodel_set-0-calling_code").val())) {
            var ddate = $("#id_tcmodel_set-0-followup_date input").val();
            if(ddate == "")
            {
                $(".field-followup_date").addClass('has-error');
                flag = false;
            }
            else
            {
                $(".field-followup_date").removeClass('has-error');
            }
        }

        var calling_remarks = $("#id_tcmodel_set-0-calling_remarks").val();
        if(calling_remarks == "")
        {
            $(".field-calling_remarks").addClass('has-error');
            flag = false;
        }
        else
        {
            $(".field-calling_remarks").removeClass('has-error');
        }
        return flag;
    });
});
