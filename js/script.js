jQuery(document).ready( function($) {
	var url_str = document.location.href;
	console.log(url_str);
	jQuery('.check-column>input[type="checkbox"]').on('change', function(){
		//alert(url_str);
		if (~url_str.indexOf("primer_slug")) {
			var thischeck = $(this).val();
			console.log(thischeck);
			jQuery('input[type="checkbox"]').each(function(){
				if ($(this).val() != thischeck){
					$(this).prop('checked', false);
					$('select[name="' + $(this).val() + 'expformoption"]').prop('selectedIndex',0);
				}
			});
		};
	});
}); 
