import { useEffect, useState } from 'react';
import api from '../api/client';
import PageShell from '../components/layout/PageShell';
import { Badge } from '../components/ui/badge';
import { Table, TableCell, TableHead, TableRow } from '../components/ui/table';

const statuses = ['saved', 'applied', 'OA', 'interview', 'rejected', 'offer'];

const Tracker = () => {
  const [applications, setApplications] = useState([]);

  const loadApplications = async () => {
    const response = await api.get('/applications');
    setApplications(response.data);
  };

  useEffect(() => {
    loadApplications();
  }, []);

  const updateStatus = async (id, status) => {
    await api.patch(`/applications/${id}/status`, { status });
    loadApplications();
  };

  return (
    <PageShell title="Application Tracker" subtitle="Track application stages with real-time updates.">
      <Table>
        <thead>
          <tr>
            <TableHead>Role</TableHead>
            <TableHead>Company</TableHead>
            <TableHead>Location</TableHead>
            <TableHead>Status</TableHead>
          </tr>
        </thead>
        <tbody>
          {applications.map((app) => (
            <TableRow key={app.id}>
              <TableCell className="font-medium text-white">{app.title}</TableCell>
              <TableCell>{app.company}</TableCell>
              <TableCell>{app.location}</TableCell>
              <TableCell>
                <div className="flex items-center gap-3">
                  <Badge className="bg-white/10">{app.status}</Badge>
                  <select
                    className="glass-input max-w-[140px]"
                    value={app.status}
                    onChange={(e) => updateStatus(app.id, e.target.value)}
                  >
                    {statuses.map((status) => (
                      <option key={status} value={status}>
                        {status}
                      </option>
                    ))}
                  </select>
                </div>
              </TableCell>
            </TableRow>
          ))}
        </tbody>
      </Table>
    </PageShell>
  );
};

export default Tracker;
