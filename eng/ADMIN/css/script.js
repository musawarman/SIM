function set_action(param, act)
{
  param.form.act.value = act;
  param.form.target = '_self';
  param.form.submit();
}

function select_all(param)
{
    form_length = param.form.length
    for (i=0;i<form_length;i++)
    {
        if (param.form.item(i).type == 'checkbox')
        {
            param.form.item(i).checked = true
        }   
    }
}

function check_max_input(param, index, max_length)
{
    if (param.form(index).value.length > max_length)
    {
	    param.form(index).value = param.form(index).value.substring(0,max_length);
    }
}

function jump_url(url)
{ 
    document.location = url;
}