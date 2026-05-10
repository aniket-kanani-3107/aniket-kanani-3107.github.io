import { Outlet } from 'react-router-dom';
import Sidebar from './components/layout/Sidebar';
import Topbar from './components/layout/Topbar';

const App = () => (
  <div className="min-h-screen bg-gradient-to-br from-base-950 via-base-900 to-base-800">
    <div className="mx-auto flex max-w-7xl flex-col gap-6 px-4 py-6 lg:grid lg:grid-cols-[260px_1fr]">
      <Sidebar />
      <main className="flex flex-col gap-6">
        <Topbar />
        <Outlet />
      </main>
    </div>
  </div>
);

export default App;
