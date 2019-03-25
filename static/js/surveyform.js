//Submit Survey
$(document)
  .ready(function () {
    $('.ui.form')
      .form({
        fields: {
          name: {
            identifier: 'name',
            rules: [
              {
                type: 'empty',
                prompt: 'Please enter your name'
              }
            ]
          },
          email: {
            identifier: 'email',
            rules: [
              {
                type: 'empty',
                prompt: 'Please enter your e-mail'
              },
              {
                type: 'email',
                prompt: 'Please enter a valid e-mail'
              }
            ]
          },
          age: {
            identifier: 'age',
            rules: [
              {
                type: 'empty',
                prompt: 'Please enter your age'
              }
            ]
          },
        }
      })
      ;
  })
  ;
//Dropdown Menus
$('.dropdown')
  .dropdown({
    // you can use any ui transition
    transition: 'drop'
  })
  ;
//Checkboxes
$('.ui.checkbox')
  .checkbox()
  ;
document.addEventListener('DOMContentLoaded', () => {
  // Change the visibility of error messages
  document.querySelector('.ui.submit.button').onclick = function () {
    document.querySelector('.ui.error.message').style.visibility = 'visible';
  };
});