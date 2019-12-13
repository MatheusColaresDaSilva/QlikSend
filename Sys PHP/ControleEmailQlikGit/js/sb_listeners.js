(function($) {
$(document).ready(function() {
     jQuery.ajax({
      type: 'POST',
      url: './src/funcoes.php',
      data: {opcao: 'folder'},

      success: function (obj) {
         $('#nameapp').html('<option value="" selected="selected">Selecione um Aplicação</option>'+ obj).show();
       }
    })
    
	$('#nameapp').change(function(e){

      if($(this).val() == ""){
        $('#btnadd').prop('disabled', true);
        $('#btnremove').prop('disabled', true);
      } 
     

      jQuery.ajax({
        type: 'POST',
        url: './src/funcoes.php',
        data: {opcao: 'report', file: $('#nameapp').val() },

        success: function (obj) {
                 $('#namereport').html('<option value="" selected="selected">Selecione um Relatório</option>'+obj).show();
                }
      });
      
      jQuery.ajax({
        type: 'POST',
        url: './src/funcoes.php',
        data: {opcao: 'emailapp', file: $('#nameapp').val() },

        success: function (obj) {

                 $('#emailreport').html(obj).show();
                }
      });
    })

   $('#namereport').change(function(e){

      if($(this).val() == ""){
        $('#btnadd').prop('disabled', true);
        $('#btnremove').prop('disabled', true);
      } 
      else{
        $('#btnadd').prop('disabled', false);
        $('#btnremove').prop('disabled', false);
      }
      

      jQuery.ajax({
        type: 'POST',
        url: './src/funcoes.php',
        data:{opcao: 'emailreport' , file: $('#nameapp').val(),  report: $('#namereport').val()},
    
        success: function(obj){

          $('#emailreport').html(obj).show();
        }
      })

    })

    $('#btnadd').click(function(e){

     var emails = new Array();
     var stremails;
     
     $('#emailreport option').each(function(){
        
        emails.push($(this).val());

     }); 
     
     var email = prompt("Digite o Email","@santacasamaringa.com.br");
     
     if(email === null){
      return;
     }
     else if (email == null || email == "@santacasamaringa.com.br") {
        alert("Digite um email Válido");
      }
     else {

        emails.push(email);
        stremails = emails.join([separador = ';']) +';';

        jQuery.ajax({
          type:'POST',
          url:'./src/funcoes.php',
          data: {opcao: 'modifyEmail', emailsend: stremails, file: $('#nameapp').val(),  report: $('#namereport').val()},

          success: function(obj){

            $('#namereport').change();
          }

        })

      }
    
    })

   $('#btnremove').click(function(e){

     if($('#emailreport').val() == null){
       alert('Selecione um email da lista');
     } 
     else{
       var emails = new Array();
       var stremails;
       
       $('#emailreport option').each(function(){
          
          if( !($(this).val() == $('#emailreport').val())){

            emails.push($(this).val()); 
          }
          
       });
       
       if (confirm("Confirma remover esse Email!")) {
         
          stremails = emails.join([separador = ';']) +';';

          jQuery.ajax({
            type:'POST',
            url:'./src/funcoes.php',
            data: {opcao: 'modifyEmail', emailsend: stremails, file: $('#nameapp').val(),  report: $('#namereport').val()},

            success: function(obj){

              $('#namereport').change();
            }

          })
       } 
     }
     
    })

  })
}) (jQuery); 