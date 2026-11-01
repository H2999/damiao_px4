# Experimental rotational DOB diagnostics. Does not indicate successful recovery.
uint64 timestamp
uint64 timestamp_sample
uint32 event_count
uint8 IDLE = 0
uint8 SHOCK = 1
uint8 RECOVERY = 2
uint8 state
uint8 mode # 0 disabled, 1 monitor only, 2 compensation permitted
bool imu_fresh
bool allocator_ok
bool compensation_active
float32 force_norm # m/s^2, specific force (includes support against gravity)
float32 rate_norm # rad/s
float32[3] disturbance # estimated equivalent normalized torque disturbance
float32[3] correction # applied normalized torque correction, before battery scaling
