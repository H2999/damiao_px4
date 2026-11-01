# Per-EKF delayed-time adaptive process noise diagnostics.
uint64 timestamp
uint64 timestamp_sample
uint32 accel_device_id
uint32 gyro_device_id
uint32 event_count
uint8 state # 0 idle, 1 shock, 2 recovery window
bool enabled
float32 variance_scale # multiplier on nominal IMU noise variances, not standard deviations
