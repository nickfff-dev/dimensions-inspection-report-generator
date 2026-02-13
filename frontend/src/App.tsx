import InspectionForm from './components/InspectionForm';
import './App.css';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>📊 Excel Generation Tool</h1>
        <p>Generate dimensional inspection reports from your form data</p>
      </header>
      <main className="App-main">
        <InspectionForm />
      </main>
    </div>
  );
}

export default App;