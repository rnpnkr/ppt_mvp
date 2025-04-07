import React, { useState } from 'react';
import axios from 'axios';
import './App.css';

function App() {
  const [file, setFile] = useState(null);

  const handleFileChange = (event) => {
    setFile(event.target.files[0]);
  };

  const handleUpload = async () => {
    if (!file) return;
    const formData = new FormData();
    formData.append('file', file);
    try {
      const response = await axios.post('http://127.0.0.1:8000/upload', formData);
      console.log(response.data);
    } catch (error) {
      console.error('Upload failed:', error);
    }
  };

  return (
    <div className="App">
      <h1>PPT MVP</h1>
      <input type="file" accept=".pptx" onChange={handleFileChange} />
      <button onClick={handleUpload}>Upload Template</button>
    </div>
  );
}

export default App;