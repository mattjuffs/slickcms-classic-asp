function Email() {
    // buils up an email address to prevent spam
	var user = 'webmaster';
	var domain = 'slickhouse';
	var extension = '.com';
	var subject = 'Contact from website';
	var email = user + '@' + domain + extension
	document.write('<a href=\"mailto:' + email + '?subject=' + subject + '\">' + email +'</a>');
}

function GenerateURL(title){
	// used as a helper to generate a URL automatically
	title = title.toLowerCase();
	title = title.replace('/','');
	title = title.replace(new RegExp(/\\/g),'');
	title = title.replace(/ /g,'-');
	if (document.getElementById("url").value == ''){
		document.getElementById("url").value = title;
	}
}

function Toggle(strObj) {
    // shows/hides an element
    if (document.getElementById(strObj).style.display == 'block') {
        document.getElementById(strObj).style.display = 'none'
    } else {
        document.getElementById(strObj).style.display = 'block'
    }
}

function ValidateComment() {
    // validates a visitor's comment
    var valid = true;
    var error = '';
    var commentError = document.getElementById('comments-error');

    // required fields
    if (document.getElementById('name').value == '') {
        error += '<em>Name</em><br />';
        valid = false;
    }

    if (document.getElementById('email').value == '') {
        error += '<em>Email</em><br />';
        valid = false;
    }

    if (document.getElementById('comment').value == '') {
        error += '<em>Comment</em><br />';
        valid = false;
    }
    
    if (valid == false){
        error = 'Please ensure all of the following are filled in:<br />' + error;
        commentError.style.display = 'block';
        commentError.innerHTML = error;
    } else {
        commentError.style.display = 'none';
    }

    return valid;
}