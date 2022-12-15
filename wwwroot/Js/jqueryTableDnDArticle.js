$(document).ready(function() {
	// Initialise the first table (as before)
	
	$('#table-3').tableDnD({
		   
	    onDrop: function(table, row) {
	$('#siralama').val('')
	var rows = table.tBodies[0].rows;
	var siralama = $('#siralama').val();
            for (var i=0; i<rows.length; i++) {
		siralama += rows[i].id+"-";
            }
	   $('#siralama').val(siralama);
	   function formyolla(){
$.ajax({
   type: 'POST',
   url: 'save.asp',
   data: $('#siralama').serialize()

});
}
	eval(formyolla())	    
        }
	}); 
	
	
    $("#table-3 tr").hover(function() {
          $(this.cells[0]).addClass('showDragHandle');
    }, function() {
          $(this.cells[0]).removeClass('showDragHandle');
    });
    
});
