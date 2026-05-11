import { useEffect, useState } from 'react';
import { motion } from 'framer-motion';
import api from '../api/client';
import { Card, CardContent, CardHeader, CardTitle } from '../components/ui/card';
import { Badge } from '../components/ui/badge';
import { Table, TableCell, TableHead, TableRow } from '../components/ui/table';

const StatCard = ({ label, value, accent }) => (
  <motion.div whileHover={{ y: -4 }} transition={{ duration: 0.2 }}>
    <Card>
      <CardHeader>
        <CardTitle>{label}</CardTitle>
        <Badge className={accent}>{value}</Badge>
      </CardHeader>
      <CardContent>
        <p className="text-xs text-slate-400">Last updated just now</p>
      </CardContent>
    </Card>
  </motion.div>
);

const Dashboard = () => {
  const [data, setData] = useState({
    metrics: { totalJobs: 0, appliedJobs: 0, avgAts: 0 },
    latestJobs: [],
    recentResumes: [],
    activity: [],
  });

  useEffect(() => {
    const fetchDashboard = async () => {
      const response = await api.get('/dashboard');
      setData(response.data);
    };

    fetchDashboard();
  }, []);

  const trend = data.latestJobs
    .map((job) => job.ats_score || 0)
    .slice(0, 6)
    .reverse();

  return (
    <div className="flex flex-col gap-6">
      <div className="grid gap-4 md:grid-cols-3">
        <StatCard label="Total Jobs Found" value={data.metrics.totalJobs} accent="bg-emerald-500/20 text-emerald-200" />
        <StatCard label="Jobs Applied" value={data.metrics.appliedJobs} accent="bg-indigo-500/20 text-indigo-200" />
        <StatCard label="ATS Score Avg" value={`${data.metrics.avgAts}%`} accent="bg-pink-500/20 text-pink-200" />
      </div>

      <div className="grid gap-6 lg:grid-cols-[1.2fr_0.8fr]">
        <Card>
          <CardHeader>
            <CardTitle>Latest Jobs</CardTitle>
            <Badge>{data.latestJobs.length} new</Badge>
          </CardHeader>
          <CardContent>
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
                {data.latestJobs.map((job) => (
                  <TableRow key={job.id}>
                    <TableCell className="font-medium text-white">{job.title}</TableCell>
                    <TableCell>{job.company}</TableCell>
                    <TableCell>{job.location}</TableCell>
                    <TableCell>
                      <Badge className="bg-white/10">{job.status}</Badge>
                    </TableCell>
                  </TableRow>
                ))}
              </tbody>
            </Table>
          </CardContent>
        </Card>

        <div className="flex flex-col gap-6">
          <Card>
            <CardHeader>
              <CardTitle>ATS Score Trend</CardTitle>
              <Badge className="bg-white/10">Last {trend.length} jobs</Badge>
            </CardHeader>
            <CardContent>
              <div className="flex items-end gap-3">
                {trend.map((value, index) => (
                  <div key={`${value}-${index}`} className="flex flex-1 flex-col items-center gap-2">
                    <div className="flex h-28 w-full items-end rounded-full bg-white/10">
                      <div
                        className="w-full rounded-full bg-gradient-to-t from-accent to-accentSoft"
                        style={{ height: `${Math.max(value, 8)}%` }}
                      />
                    </div>
                    <span className="text-xs text-slate-400">{value}%</span>
                  </div>
                ))}
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Recent Resumes</CardTitle>
              <Badge className="bg-white/10">{data.recentResumes.length}</Badge>
            </CardHeader>
            <CardContent className="flex flex-col gap-3">
              {data.recentResumes.map((resume) => (
                <div key={resume.id} className="rounded-xl border border-white/10 bg-white/5 p-3">
                  <p className="text-sm font-medium text-white">{resume.version_name}</p>
                  <p className="text-xs text-slate-400">{resume.created_at}</p>
                </div>
              ))}
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Activity Feed</CardTitle>
            </CardHeader>
            <CardContent className="flex flex-col gap-3">
              {data.activity.map((event) => (
                <div key={event.id} className="rounded-xl border border-white/10 bg-white/5 p-3">
                  <p className="text-xs uppercase text-slate-400">{event.type}</p>
                  <p className="text-sm text-white">{event.message}</p>
                  <p className="text-xs text-slate-500">{event.created_at}</p>
                </div>
              ))}
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;
