// SPDX-License-Identifier: BSD-3-Clause
// Copyright (c) 2026 PX4 Development Team.
#pragma once

#include <cmath>
#include <cstdint>

namespace collision_recovery
{

// Time is supplied by the caller: delayed IMU time in EKF, sensor time in control.
// This is an event window, not a claim that attitude or translation is recovered.
class ImpactDetector
{
public:
	// 正常 碰撞 恢复
	enum class State : uint8_t { Idle = 0, Shock = 1, Recovery = 2 };

	void reset() { _last_sample = 0; _last_impact = 0; _events = 0; _state = State::Idle; }

	void update(uint64_t time_us, float force_norm, float rate_norm, bool clipped,
		    bool enabled, float force_threshold, float rate_threshold)
	{
		if (!enabled) {
			reset();
			return;
		}

		if (time_us == 0 || time_us == _last_sample) { return; }

		if (time_us < _last_sample) { reset(); }

		_last_sample = time_us;
		_state = stateAt(time_us);
		const bool valid = std::isfinite(force_norm) && std::isfinite(rate_norm);
		const bool impact = clipped || (valid && (force_norm > force_threshold || rate_norm > rate_threshold));

		if (impact) {
			if (_state == State::Idle) { ++_events; }

			_last_impact = time_us;
			_state = State::Shock;
		}
	}

	State stateAt(uint64_t time_us) const
	{
		if (_last_impact == 0 || time_us < _last_impact) { return State::Idle; }

		const uint64_t elapsed = time_us - _last_impact;
		return elapsed < shock_us ? State::Shock : (elapsed < window_us ? State::Recovery : State::Idle);
	}

	float varianceScale(uint64_t time_us, float maximum) const
	{
		maximum = std::isfinite(maximum) ? fminf(fmaxf(maximum, 1.f), 100.f) : 1.f;
		const State state = stateAt(time_us);

		if (state == State::Shock) { return maximum; }

		if (state == State::Recovery) {
			const float remaining = float(window_us - (time_us - _last_impact)) / float(window_us - shock_us);
			return 1.f + (maximum - 1.f) * remaining;
		}

		return 1.f;
	}

	State state() const { return _state; }
	uint32_t events() const { return _events; }
	static constexpr uint64_t shock_us = 80000;
	static constexpr uint64_t window_us = 580000;

private:

	uint64_t _last_sample{0};
	uint64_t _last_impact{0};
	uint32_t _events{0};
	State _state{State::Idle};
};

} // namespace collision_recovery
