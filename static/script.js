// script.js

// Function to show the loading overlay and progress bar
// function showLoadingOverlay() {
//   const overlay = document.createElement("div");
//   overlay.classList.add("overlay");

//   const progressBar = document.createElement("div");
//   progressBar.classList.add("progress-bar");

//   overlay.appendChild(progressBar);
//   document.body.appendChild(overlay);
  
// }

// // // Function to hide the loading overlay
// function hideLoadingOverlay() {
//   const overlay = document.querySelector(".overlay");
//   if (overlay) {
//     document.body.removeChild(overlay);
//   }
// }


// document.getElementById("uploadForm").addEventListener("submit", async function (event) {
//   event.preventDefault();

//   showLoadingOverlay();
  

//   const file1 = document.getElementById("file1").files[0];
//   console.log(file1);

//   // Use FormData to send the file in the request
//   const formData = new FormData();
//   formData.append('file1', file1);

//   // Get column names of File 1 using AJAX request
//   const response = await fetch('/get_column_names', {
//     method: 'POST',
//     body: formData,
//   })
//   const data = await response.json();
//   console.log(data);

//   const checkboxesDiv = document.getElementById("checkboxContainer");
//   checkboxesDiv.innerHTML = '';

//   // Create checkboxes based on column names of File 1
//   data.column_names.forEach(columnName => {
//     const checkbox = document.createElement("input");
//     checkbox.type = "checkbox";
//     checkbox.name = "selected_columns";
//     checkbox.value = columnName;

//     const label = document.createElement("label");
//     label.innerHTML = columnName;

//     checkboxesDiv.appendChild(checkbox);
//     checkboxesDiv.appendChild(label);
//     checkboxesDiv.appendChild(document.createElement("br"));
//   });

//   hideLoadingOverlay();
// });


// script.js

function showLoadingOverlay() {
  const overlay = document.createElement("div");
  overlay.classList.add("overlay");

  const loader = document.createElement("div");
  loader.classList.add("loader");

  overlay.appendChild(loader);
  document.body.appendChild(overlay);
}

// Function to hide the loading overlay
function hideLoadingOverlay() {
  const overlay = document.querySelector(".overlay");
  if (overlay) {
    document.body.removeChild(overlay);
  }
}


// Function to request wake lock
async function requestWakeLock() {
  try {
    // Check if the Wake Lock API is supported by the browser
    if ('wakeLock' in navigator) {
      const wakeLock = await navigator.wakeLock.request('screen');
      console.log('Wake lock active:', wakeLock);
    } else {
      console.log('Wake Lock API is not supported.');
    }
  } catch (error) {
    console.error('Could not request wake lock:', error);
  }
}

// Function to release wake lock
async function releaseWakeLock() {
  try {
    if ('wakeLock' in navigator) {
      const wakeLock = await navigator.wakeLock.release();
      console.log('Wake lock released:', wakeLock);
    }
  } catch (error) {
    console.error('Could not release wake lock:', error);
  }
}


async function uploadFile(formData) {
  try {
    const response = await fetch('/upload', {
      method: 'POST',
      body: formData,
    });

    if (response.ok) {
      // If the response is successful, navigate to the next page
      window.location.href = "/process";
    } else {
      alert("File upload failed!");
    }
  } catch (error) {
    console.error("Error during file upload:", error);
    alert("File upload failed!");
  } finally {
    hideLoadingOverlay();
  }
}


document.getElementById("uploadForm").addEventListener("submit", async function(event) {
  event.preventDefault();

  // Request wake lock when the file upload starts
  await requestWakeLock();

  showLoadingOverlay();

  const file1 = document.getElementById("file1").files[0];
  const file2 = document.getElementById("file2").files[0];

  const formData = new FormData();
  formData.append('file1', file1);
  formData.append('file2', file2);

  uploadFile(formData);

  // Request wake lock when the file upload starts
  await releaseWakeLock();
});
