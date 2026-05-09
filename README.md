# Academic Block Simulation Backend

Node.js + Express backend for uploading timetable Excel files and generating room utilization metrics.

## Run

```bash
npm install
npm run dev
```

## API

- `POST /upload` with multipart form-data field `file`
- `GET /metrics`
- `GET /health`

## Excel sheets

- `Rooms`: `room_id`, `capacity`, `type`, `floor`
- `Timetable`: `class_id`, `subject`, `room_id`, `start_time`, `end_time`, `day`
- `Enrollment`: `class_id`, `student_count`

## Notes

- Upload processing uses in-memory storage and keeps the latest computed metrics in memory for the `GET /metrics` endpoint.
- Utilization is calculated from used room time over the inferred available timetable window across uploaded days.
