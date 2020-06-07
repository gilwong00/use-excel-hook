import React from 'react';
import useExcel from './hooks/useExcel';
import './App.css';

const data = [{ name: 'test', age: 29 }];

// header example
const headers = [
  { header: 'Name', key: 'name' },
  { header: 'Age', key: 'age' }
];

function App() {
  const { exportToExcel } = useExcel();

  return (
    <div className='App'>
      <button onClick={() => exportToExcel({ data })}>
        Export to excel(basic)
      </button>
      <button onClick={() => exportToExcel({ headers, data })}>
        Export to excel(Custom headers)
      </button>
    </div>
  );
}

export default App;
