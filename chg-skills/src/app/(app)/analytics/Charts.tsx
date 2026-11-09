"use client";
import { Bar, BarChart, CartesianGrid, Legend, Line, ComposedChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { Card, CardHeader } from "@/components/ui";

const NAVY = "#14284B", BLUE = "#1E63C4", GREEN = "#2E8B57", GRID = "#E3E7EC", INK = "#5D6B7E";
const monthLabel = (m: string) => new Date(`${m}-15`).toLocaleDateString("en-US", { month: "short", year: "2-digit" });

export function Charts({ byArea, dist, trend }: {
  byArea: { name: string; assigned: number; assignedQualified: number; qualified: number; gaps: number }[];
  dist: number[];
  trend: { month: string; qualificationsVerified: number; trainingCompleted: number; activeQualifications: number }[];
}) {
  const crossData = [{ k: "0 areas", v: dist[0] }, { k: "1 area", v: dist[1] }, { k: "2 areas", v: dist[2] }, { k: "3+ areas", v: dist[3] }];
  return (
    <div className="grid gap-5 xl:grid-cols-2">
      <Card>
        <CardHeader title="Qualification coverage by department" subtitle="People qualified (any shift) vs people assigned" />
        <div className="h-80 p-3">
          <ResponsiveContainer>
            <BarChart data={byArea} margin={{ left: -10, right: 10, bottom: 50 }}>
              <CartesianGrid stroke={GRID} vertical={false} />
              <XAxis dataKey="name" angle={-40} textAnchor="end" interval={0} tick={{ fontSize: 11, fill: INK }} height={70} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: INK }} />
              <Tooltip />
              <Legend verticalAlign="top" wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="assigned" name="Assigned" fill="#AFC1DC" radius={[3, 3, 0, 0]} />
              <Bar dataKey="qualified" name="Qualified" fill={BLUE} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>
      <Card>
        <CardHeader title="Cross-training progress" subtitle="Employees by number of production areas qualified" />
        <div className="h-80 p-3">
          <ResponsiveContainer>
            <BarChart data={crossData} margin={{ left: -10, right: 10 }}>
              <CartesianGrid stroke={GRID} vertical={false} />
              <XAxis dataKey="k" tick={{ fontSize: 12, fill: INK }} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: INK }} />
              <Tooltip />
              <Bar dataKey="v" name="Employees" fill={GREEN} radius={[3, 3, 0, 0]} label={{ position: "top", fontSize: 12, fill: INK }} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>
      <Card className="xl:col-span-2">
        <CardHeader title="Training completion and qualification growth" subtitle="Last 12 months" />
        <div className="h-80 p-3">
          <ResponsiveContainer>
            <ComposedChart data={trend.map((t) => ({ ...t, label: monthLabel(t.month) }))} margin={{ left: -10, right: 10 }}>
              <CartesianGrid stroke={GRID} vertical={false} />
              <XAxis dataKey="label" tick={{ fontSize: 11, fill: INK }} />
              <YAxis yAxisId="l" allowDecimals={false} tick={{ fontSize: 11, fill: INK }} />
              <YAxis yAxisId="r" orientation="right" allowDecimals={false} tick={{ fontSize: 11, fill: INK }} />
              <Tooltip />
              <Legend verticalAlign="top" wrapperStyle={{ fontSize: 12 }} />
              <Bar yAxisId="l" dataKey="qualificationsVerified" name="Qualifications verified" fill={BLUE} radius={[3, 3, 0, 0]} />
              <Bar yAxisId="l" dataKey="trainingCompleted" name="Training completed" fill="#F2C230" radius={[3, 3, 0, 0]} />
              <Line yAxisId="r" dataKey="activeQualifications" name="Active qualifications (month end)" stroke={NAVY} strokeWidth={2.5} dot={{ r: 3 }} />
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      </Card>
    </div>
  );
}
