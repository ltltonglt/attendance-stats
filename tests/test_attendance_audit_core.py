from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from attendance_audit_core import (
    AttendanceRecord,
    audit_person_records,
    audit_personnel,
    judge_plate_consistency,
    judge_product_consistency,
    parse_attendance_time,
)


class AttendanceAuditCoreTests(unittest.TestCase):
    def record(
        self,
        name: str = "张三",
        email: str = "zhangsan@example.com",
        time: datetime = datetime(2026, 5, 13, 8, 0, 0),
        product: str = "5G NR",
        scene: str = "无",
        plate: str = "",
    ) -> AttendanceRecord:
        return AttendanceRecord(
            name=name,
            email=email,
            attendance_time=time,
            product=product,
            car_scene=scene,
            plate_number=plate,
        )

    def test_parse_attendance_time_accepts_string_and_datetime(self) -> None:
        value = datetime(2026, 5, 13, 8, 15, 31)
        self.assertEqual(parse_attendance_time(value), value)
        self.assertEqual(parse_attendance_time("2026/05/13 08:15:31"), value)
        self.assertEqual(
            parse_attendance_time("2026/01/23 18:00:00-2026/07/01 08:25:00"),
            datetime(2026, 7, 1, 8, 25, 0),
        )

    def test_eight_hour_boundary_is_yes(self) -> None:
        result = audit_person_records(
            "5G NR网络优化",
            [
                self.record(time=datetime(2026, 5, 13, 8, 0, 0), product="5G NR"),
                self.record(time=datetime(2026, 5, 13, 16, 0, 0), product="5G NR"),
            ],
        )
        self.assertEqual(result.meets_8_hours, "是")
        self.assertEqual(result.product_consistency, "一致")
        self.assertEqual(result.remark, "")

    def test_uses_car_scene_times_when_normal_records_less_than_two(self) -> None:
        result = audit_person_records(
            "5G NR网络优化",
            [
                self.record(time=datetime(2026, 5, 13, 12, 0, 0), product="5G NR", scene="无"),
                self.record(time=datetime(2026, 5, 13, 8, 0, 0), product="5G NR", scene="开始用车", plate="川A12345"),
                self.record(time=datetime(2026, 5, 13, 18, 0, 0), product="5G NR", scene="结束用车", plate="川A12345"),
            ],
        )
        self.assertEqual(result.sign_in_time, datetime(2026, 5, 13, 8, 0, 0))
        self.assertEqual(result.sign_out_time, datetime(2026, 5, 13, 18, 0, 0))
        self.assertEqual(result.product_consistency, "一致")
        self.assertEqual(result.plate_consistency, "一致")
        self.assertEqual(result.remark, "签到时间使用开始用车时间，签退时间使用结束用车时间")

    def test_multiple_start_car_records_match_any_end_plate(self) -> None:
        result = audit_person_records(
            "LTE网络优化",
            [
                self.record(time=datetime(2026, 5, 13, 8, 18, 57), product="LTE", scene="无"),
                self.record(time=datetime(2026, 5, 13, 8, 56, 58), product="LTE", scene="开始用车", plate="川ACX5359"),
                self.record(time=datetime(2026, 5, 13, 10, 40, 8), product="LTE", scene="开始用车", plate="川AC75307"),
                self.record(time=datetime(2026, 5, 13, 15, 50, 19), product="LTE", scene="结束用车", plate="川AC75307"),
            ],
        )
        self.assertEqual(result.sign_in_plate, "川AC75307")
        self.assertEqual(result.sign_out_plate, "川AC75307")
        self.assertEqual(result.plate_consistency, "一致")
        self.assertEqual(result.remark, "签退时间使用结束用车时间，存在两条开始用车")

    def test_normal_records_keep_priority_when_two_or_more_exist(self) -> None:
        result = audit_person_records(
            "LTE网络优化",
            [
                self.record(time=datetime(2026, 5, 13, 9, 0, 0), product="LTE", scene="无"),
                self.record(time=datetime(2026, 5, 13, 17, 0, 0), product="LTE", scene="无"),
                self.record(time=datetime(2026, 5, 13, 8, 0, 0), product="5G NR", scene="开始用车", plate="川A12345"),
                self.record(time=datetime(2026, 5, 13, 18, 0, 0), product="5G NR", scene="结束用车", plate="川A12345"),
            ],
        )
        self.assertEqual(result.sign_in_time, datetime(2026, 5, 13, 9, 0, 0))
        self.assertEqual(result.sign_out_time, datetime(2026, 5, 13, 17, 0, 0))
        self.assertEqual(result.product_consistency, "一致")
        self.assertEqual(result.remark, "")

    def test_product_consistency_all_cases(self) -> None:
        good = self.record(product="LTE")
        bad = self.record(product="5G NR")

        self.assertEqual(judge_product_consistency("LTE网络优化", good, good), "一致")
        self.assertEqual(judge_product_consistency("LTE网络优化", bad, good), "签到制式不一致")
        self.assertEqual(judge_product_consistency("LTE网络优化", good, bad), "签退制式不一致")
        self.assertEqual(judge_product_consistency("LTE网络优化", bad, bad), "签到和签退制式不一致")

    def test_plate_consistency_all_cases(self) -> None:
        self.assertEqual(judge_plate_consistency("川A12345", "川A12345"), "一致")
        self.assertEqual(judge_plate_consistency("川A12345", "川B12345"), "不一致")
        self.assertEqual(judge_plate_consistency("", "川B12345"), "缺少签到")
        self.assertEqual(judge_plate_consistency("川A12345", ""), "缺少签退")
        self.assertEqual(judge_plate_consistency("", ""), "未用车")

    def test_duplicate_name_uses_name_and_email_when_exporting(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            person_file = root / "人员信息.xlsx"
            attendance_dir = root / "打卡数据导出"
            output_file = root / "结果.xlsx"
            attendance_dir.mkdir()

            person_wb = Workbook()
            person_ws = person_wb.active
            person_ws.append(["项目号", "姓名", "用户邮箱", "打卡制式", "工号（电旗）", "姓名（电旗）"])
            person_ws.append(["P1", "王杰", "first@example.com", "5G NR网络优化", "001", "王杰.1"])
            person_ws.append(["P2", "王杰", "second@example.com", "LTE网络优化", "002", "王杰.2"])
            person_ws.append(["P3", "李四", "lisi@example.com", "LTE网络优化", "003", "李四"])
            person_wb.save(person_file)

            attendance_wb = Workbook()
            attendance_ws = attendance_wb.active
            attendance_ws.title = "data"
            attendance_ws.append(["考勤类型", "姓名", "用户邮箱", "考勤时间", "主产品", "用车场景", "车牌号"])
            attendance_ws.append(["签到", "王杰", "first@example.com", "2026/05/13 08:00:00", "5G NR", "无", ""])
            attendance_ws.append(["签到", "王杰", "first@example.com", "2026/05/13 17:00:00", "5G NR", "无", ""])
            attendance_ws.append(["签到", "王杰", "second@example.com", "2026/05/13 09:00:00", "LTE", "无", ""])
            attendance_ws.append(["签到", "王杰", "second@example.com", "2026/05/13 16:00:00", "LTE", "无", ""])
            attendance_ws.append(["签到", "李四", "wrong@example.com", "2026/05/13 08:30:00", "LTE", "无", ""])
            attendance_ws.append(["签到", "李四", "another@example.com", "2026/05/13 18:00:00", "LTE", "无", ""])
            attendance_ws.append(["签到", "李四", "another@example.com", "2026/05/13 07:50:00", "LTE", "开始用车", "川A12345"])
            attendance_ws.append(["签到", "李四", "another@example.com", "2026/05/13 18:05:00", "LTE", "结束用车", "川A12345"])
            attendance_ws.append(["请假", "李四", "another@example.com", "2026/01/23 18:00:00-2026/07/01 08:25:00", "LTE", "无", ""])
            attendance_wb.save(attendance_dir / "SIGN.xlsx")

            summary = audit_personnel(person_file, attendance_dir, output_file)
            self.assertEqual(summary.person_count, 3)
            self.assertEqual(summary.attendance_file_count, 1)
            self.assertEqual(summary.attendance_record_count, 8)
            self.assertEqual(summary.duplicate_names, ["王杰"])
            self.assertIsNone(summary.merged_output_path)

            output_wb = load_workbook(output_file, data_only=True)
            output_ws = output_wb.active

            self.assertEqual(output_ws.cell(2, 7).value, datetime(2026, 5, 13, 8, 0, 0))
            self.assertEqual(output_ws.cell(2, 8).value, datetime(2026, 5, 13, 17, 0, 0))
            self.assertEqual(output_ws.cell(2, 9).value, "是")
            self.assertEqual(output_ws.cell(2, 10).value, "一致")
            self.assertEqual(output_ws.cell(2, 11).value, "/")
            self.assertEqual(output_ws.cell(2, 12).value, "/")
            self.assertEqual(output_ws.cell(2, 13).value, "未用车")
            self.assertEqual(output_ws.cell(2, 14).value, None)

            self.assertEqual(output_ws.cell(3, 7).value, datetime(2026, 5, 13, 9, 0, 0))
            self.assertEqual(output_ws.cell(3, 8).value, datetime(2026, 5, 13, 16, 0, 0))
            self.assertEqual(output_ws.cell(3, 9).value, "否")
            self.assertEqual(output_ws.cell(3, 10).value, "一致")

            self.assertEqual(output_ws.cell(4, 7).value, datetime(2026, 5, 13, 8, 30, 0))
            self.assertEqual(output_ws.cell(4, 8).value, datetime(2026, 5, 13, 18, 0, 0))
            self.assertEqual(output_ws.cell(4, 11).value, "川A12345")
            self.assertEqual(output_ws.cell(4, 12).value, "川A12345")
            self.assertEqual(output_ws.cell(4, 13).value, "一致")

    def test_merge_attendance_files_when_requested(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            person_file = root / "人员信息.xlsx"
            attendance_dir = root / "打卡数据导出"
            output_file = root / "结果.xlsx"
            merged_file = root / "合并打卡数据.xlsx"
            attendance_dir.mkdir()

            person_wb = Workbook()
            person_ws = person_wb.active
            person_ws.append(["项目号", "姓名", "用户邮箱", "打卡制式", "工号（电旗）", "姓名（电旗）"])
            person_ws.append(["P1", "张三", "zhangsan@example.com", "LTE网络优化", "001", "张三"])
            person_wb.save(person_file)

            for index in range(1, 3):
                attendance_wb = Workbook()
                attendance_ws = attendance_wb.active
                attendance_ws.title = "data"
                attendance_ws.append(["考勤类型", "姓名", "用户邮箱", "考勤时间", "主产品", "用车场景", "车牌号"])
                attendance_ws.append(["签到", "张三", "zhangsan@example.com", f"2026/05/13 0{index}:00:00", "LTE", "无", ""])
                attendance_ws.append(["出差", "张三", "zhangsan@example.com", "2026/01/23 18:00:00-2026/07/01 08:25:00", "LTE", "无", ""])
                attendance_wb.save(attendance_dir / f"SIGN-{index}.xlsx")

            summary = audit_personnel(
                person_file,
                attendance_dir,
                output_file,
                merge_attendance=True,
                merged_output_file=merged_file,
            )

            self.assertEqual(summary.attendance_file_count, 2)
            self.assertEqual(summary.attendance_record_count, 2)
            self.assertEqual(summary.merged_output_path, merged_file)

            merged_wb = load_workbook(merged_file, data_only=True)
            merged_ws = merged_wb.active
            self.assertEqual(merged_ws.max_row, 5)
            self.assertEqual(merged_ws.max_column, 8)
            self.assertEqual(merged_ws.cell(1, 8).value, "来源文件")
            self.assertEqual(merged_ws.cell(2, 1).value, "签到")
            self.assertEqual(merged_ws.cell(3, 1).value, "出差")


if __name__ == "__main__":
    unittest.main()
