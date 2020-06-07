import React from 'react';
import useExcel from './hooks/useExcel';
import './App.css';

function App() {
  const { exportToExcel } = useExcel();

  const handleExport = () => {
    const data = [{ name: 'test', age: 29 }];
    exportToExcel({ data });
  };

  return (
    <div className='App'>
      <button onClick={handleExport}>Export to excel</button>
    </div>
  );
}

export default App;
