"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

import {
  ResponsiveContainer,
  PieChart,
  Pie,
  Tooltip,
  Cell,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  LineChart,
  Line,
  Legend,
} from "recharts";

const COLORS = [
  "#7C3AED",
  "#06B6D4",
  "#10B981",
  "#F59E0B",
  "#EF4444",
  "#6366F1",
  "#14B8A6",
  "#EC4899",
];

type RowData = {
  사업소: string;
  계정과목: string;
  적요?: string;
  금액: number;
  월: string;
};

export default function DashboardPage() {
  const [rows, setRows] = useState<RowData[]>([]);

  const [sales, setSales] = useState<number>(500000000);

  const [targetProfitRate, setTargetProfitRate] =
    useState<number>(25);

  const handleUpload = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = (evt) => {
      const binaryStr = evt.target?.result;

      const workbook = XLSX.read(binaryStr, {
        type: "binary",
      });

      const sheetName = workbook.SheetNames[0];

      const worksheet =
        workbook.Sheets[sheetName];

      const jsonData =
        XLSX.utils.sheet_to_json<RowData>(
          worksheet
        );

      const cleaned = jsonData.map((item) => ({
        사업소: item.사업소 || "미분류",
        계정과목:
          item.계정과목 || "기타",
        적요: item.적요 || "",
        금액: Number(item.금액 || 0),
        월: item.월 || "기타",
      }));

      setRows(cleaned);
    };

    reader.readAsBinaryString(file);
  };

  // =========================
  // 사업소별 비용
  // =========================

  const siteCostData = useMemo(() => {
    const grouped: Record<string, number> =
      {};

    rows.forEach((row) => {
      if (!grouped[row.사업소]) {
        grouped[row.사업소] = 0;
      }

      grouped[row.사업소] += row.금액;
    });

    return Object.entries(grouped).map(
      ([name, value]) => ({
        name,
        value,
      })
    );
  }, [rows]);

  // =========================
  // 계정과목별 비용
  // =========================

  const categoryData = useMemo(() => {
    const grouped: Record<string, number> =
      {};

    rows.forEach((row) => {
      if (!grouped[row.계정과목]) {
        grouped[row.계정과목] = 0;
      }

      grouped[row.계정과목] += row.금액;
    });

    return Object.entries(grouped)
      .map(([name, value]) => ({
        name,
        value,
      }))
      .sort((a, b) => b.value - a.value);
  }, [rows]);

  // =========================
  // 월별 비용 추이
  // =========================

  const monthlyData = useMemo(() => {
    const grouped: Record<string, number> =
      {};

    rows.forEach((row) => {
      if (!grouped[row.월]) {
        grouped[row.월] = 0;
      }

      grouped[row.월] += row.금액;
    });

    return Object.entries(grouped).map(
      ([month, value]) => ({
        month,
        value,
      })
    );
  }, [rows]);

  // =========================
  // 총 비용
  // =========================

  const totalCost = useMemo(() => {
    return rows.reduce(
      (sum, row) => sum + row.금액,
      0
    );
  }, [rows]);

  // =========================
  // 수익률
  // =========================

  const currentProfitRate =
    ((sales - totalCost) / sales) * 100;

  // =========================
  // 목표 비용 계산
  // =========================

  const targetCost =
    sales -
    sales * (targetProfitRate / 100);

  const cutAmount =
    totalCost - targetCost;

  // =========================
  // AI 분석 로직
  // =========================

  const aiInsights = useMemo(() => {
    const result: string[] = [];

    if (rows.length === 0) {
      return [];
    }

    // 가장 비용 큰 사업소
    const topSite =
      siteCostData.sort(
        (a, b) => b.value - a.value
      )[0];

    if (topSite) {
      result.push(
        `${topSite.name} 사업소 비용 비중이 가장 높습니다.`
      );
    }

    // 가장 비용 큰 계정과목
    const topCategory = categoryData[0];

    if (topCategory) {
      const ratio =
        (
          (topCategory.value / totalCost) *
          100
        ).toFixed(1);

      result.push(
        `${topCategory.name} 비중이 전체 비용의 ${ratio}% 수준입니다.`
      );
    }

    // 목표 수익률 비교
    if (
      currentProfitRate <
      targetProfitRate
    ) {
      result.push(
        `현재 수익률이 목표 대비 ${(
          targetProfitRate -
          currentProfitRate
        ).toFixed(1)}% 부족합니다.`
      );

      result.push(
        `약 ₩${Math.abs(
          cutAmount
        ).toLocaleString()} 비용 절감이 필요합니다.`
      );
    } else {
      result.push(
        `목표 수익률을 달성 중입니다.`
      );
    }

    // 과다 계정 감지
    categoryData.forEach((item) => {
      const ratio =
        (item.value / totalCost) * 100;

      if (ratio >= 30) {
        result.push(
          `${item.name} 비용 비중이 매우 높습니다. 관리 검토가 필요합니다.`
        );
      }
    });

    // 적요 키워드 분석
    const memoText = rows
      .map((r) => r.적요)
      .join(" ");

    if (
      memoText.includes("회식") ||
      memoText.includes("식대")
    ) {
      result.push(
        `회식/식대 관련 비용 사용 빈도가 높습니다.`
      );
    }

    if (
      memoText.includes("차량") ||
      memoText.includes("주유")
    ) {
      result.push(
        `차량 유지 관련 비용 증가 여부 확인이 필요합니다.`
      );
    }

    return result;
  }, [
    rows,
    totalCost,
    currentProfitRate,
    targetProfitRate,
    cutAmount,
    categoryData,
    siteCostData,
  ]);

  // =========================
  // 비용 통제 제안
  // =========================

  const controlSuggestions =
    useMemo(() => {
      const result: string[] = [];

      if (
        currentProfitRate <
        targetProfitRate
      ) {
        result.push(
          "사업소별 월 예산 상한 설정 권장"
        );

        result.push(
          "외주비/차량유지비 우선 절감 검토"
        );

        result.push(
          "회의비 및 복리후생비 승인 프로세스 강화 필요"
        );

        result.push(
          "고정비 비중 재점검 필요"
        );

        result.push(
          "고비용 사업소 집중 관리 권장"
        );
      } else {
        result.push(
          "현재 비용 구조 안정적"
        );

        result.push(
          "현 수준 유지 권장"
        );
      }

      return result;
    }, [
      currentProfitRate,
      targetProfitRate,
    ]);

  return (
    <main className="min-h-screen bg-slate-100 p-8">

      <div className="max-w-7xl mx-auto space-y-8">

        {/* 헤더 */}

        <div className="bg-white rounded-3xl shadow p-8">

          <h1 className="text-4xl font-bold">
            AI 사업소 비용 분석 대시보드
          </h1>

          <p className="mt-3 text-gray-500">
            사업소별 비용 / 수익률 /
            AI 비용분석 자동화
          </p>

        </div>

        {/* 업로드 */}

        <div className="bg-white rounded-3xl shadow p-8">

          <h2 className="text-2xl font-bold mb-6">
            엑셀 업로드
          </h2>

          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleUpload}
            className="border p-3 rounded-xl"
          />

          <div className="mt-5 text-gray-500">

            필수 컬럼:
            <br />
            사업소 / 계정과목 / 금액 /
            월 / 적요

          </div>

        </div>

        {/* 목표 설정 */}

        <div className="grid grid-cols-2 gap-6">

          <div className="bg-white rounded-3xl shadow p-8">

            <h2 className="text-xl font-bold mb-4">
              총 매출
            </h2>

            <input
              type="number"
              value={sales}
              onChange={(e) =>
                setSales(
                  Number(e.target.value)
                )
              }
              className="w-full border p-4 rounded-xl text-2xl"
            />

          </div>

          <div className="bg-white rounded-3xl shadow p-8">

            <h2 className="text-xl font-bold mb-4">
              목표 수익률 (%)
            </h2>

            <input
              type="number"
              value={targetProfitRate}
              onChange={(e) =>
                setTargetProfitRate(
                  Number(e.target.value)
                )
              }
              className="w-full border p-4 rounded-xl text-2xl"
            />

          </div>

        </div>

        {/* KPI */}

        <div className="grid grid-cols-4 gap-6">

          <div className="bg-white rounded-3xl shadow p-6">

            <div className="text-gray-500">
              총 비용
            </div>

            <div className="text-3xl font-bold mt-3">
              ₩
              {totalCost.toLocaleString()}
            </div>

          </div>

          <div className="bg-white rounded-3xl shadow p-6">

            <div className="text-gray-500">
              총 매출
            </div>

            <div className="text-3xl font-bold mt-3">
              ₩{sales.toLocaleString()}
            </div>

          </div>

          <div className="bg-white rounded-3xl shadow p-6">

            <div className="text-gray-500">
              현재 수익률
            </div>

            <div className="text-3xl font-bold mt-3">
              {currentProfitRate.toFixed(
                1
              )}
              %
            </div>

          </div>

          <div className="bg-white rounded-3xl shadow p-6">

            <div className="text-gray-500">
              목표 절감 필요액
            </div>

            <div className="text-3xl font-bold mt-3">
              ₩
              {Math.abs(
                cutAmount
              ).toLocaleString()}
            </div>

          </div>

        </div>

        {/* 차트 */}

        <div className="grid grid-cols-2 gap-6">

          {/* 사업소별 */}

          <div className="bg-white rounded-3xl shadow p-6 h-[500px]">

            <h2 className="text-2xl font-bold mb-6">
              사업소별 비용 비중
            </h2>

            <ResponsiveContainer
              width="100%"
              height="100%"
            >

              <PieChart>

                <Pie
                  data={siteCostData}
                  dataKey="value"
                  nameKey="name"
                  outerRadius={160}
                  label
                >

                  {siteCostData.map(
                    (_, index) => (
                      <Cell
                        key={index}
                        fill={
                          COLORS[
                            index %
                              COLORS.length
                          ]
                        }
                      />
                    )
                  )}

                </Pie>

                <Tooltip />

              </PieChart>

            </ResponsiveContainer>

          </div>

          {/* 계정과목 */}

          <div className="bg-white rounded-3xl shadow p-6 h-[500px]">

            <h2 className="text-2xl font-bold mb-6">
              계정과목별 비용
            </h2>

            <ResponsiveContainer
              width="100%"
              height="100%"
            >

              <BarChart data={categoryData}>

                <CartesianGrid strokeDasharray="3 3" />

                <XAxis dataKey="name" />

                <YAxis />

                <Tooltip />

                <Legend />

                <Bar dataKey="value" />

              </BarChart>

            </ResponsiveContainer>

          </div>

        </div>

        {/* 월별 */}

        <div className="bg-white rounded-3xl shadow p-6 h-[500px]">

          <h2 className="text-2xl font-bold mb-6">
            월별 비용 추이
          </h2>

          <ResponsiveContainer
            width="100%"
            height="100%"
          >

            <LineChart data={monthlyData}>

              <CartesianGrid strokeDasharray="3 3" />

              <XAxis dataKey="month" />

              <YAxis />

              <Tooltip />

              <Legend />

              <Line
                type="monotone"
                dataKey="value"
              />

            </LineChart>

          </ResponsiveContainer>

        </div>

        {/* AI 분석 */}

        <div className="bg-white rounded-3xl shadow p-8">

          <h2 className="text-3xl font-bold mb-8">
            AI 비용 분석 결과
          </h2>

          <div className="space-y-4">

            {aiInsights.map(
              (item, index) => (
                <div
                  key={index}
                  className="p-5 rounded-2xl border bg-red-50"
                >
                  {item}
                </div>
              )
            )}

          </div>

        </div>

        {/* 비용통제 */}

        <div className="bg-white rounded-3xl shadow p-8">

          <h2 className="text-3xl font-bold mb-8">
            목표 수익률 기반 비용 통제 제안
          </h2>

          <div className="space-y-4">

            {controlSuggestions.map(
              (item, index) => (
                <div
                  key={index}
                  className="p-5 rounded-2xl border bg-blue-50"
                >
                  {item}
                </div>
              )
            )}

          </div>

        </div>

      </div>

    </main>
  );
}
