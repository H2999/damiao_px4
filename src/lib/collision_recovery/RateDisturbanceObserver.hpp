// SPDX-License-Identifier: BSD-3-Clause
// Copyright (c) 2026 PX4 Development Team.
#pragma once

#include <cmath>

namespace collision_recovery
{

// Diagonal, identified model: alpha_i = B_i * lag(u_i) + d_i.
// u is normalized torque, B is rad/s^2 per unit normalized torque, NOT inertia.
// Cross coupling, rotor gyroscopic effects and model errors enter the residual.
class RateDisturbanceObserver
{
public:

	struct Vector {
		float values[3];
		float &operator[](unsigned i) { return values[i]; }
		const float &operator[](unsigned i) const { return values[i]; }
		const float *begin() const { return values; }
		const float *end() const { return values + 3; }
	};
	void reset() { _model_command = {}; _command = {}; _disturbance = {}; _correction = {}; }
	void setCommand(const Vector &command) { _command = command; }
	const Vector &disturbance() const { return _disturbance; }
	const Vector &correction() const { return _correction; }

	// Never change the requested sign or amplify a correction when nominal demand
	// already lies beyond the normalized range.
	static float boundedCorrection(float nominal, float correction)
	{
		if (!std::isfinite(nominal) || !std::isfinite(correction)) { return 0.f; }

		return correction > 0.f ? fminf(correction, fmaxf(0.f, 1.f - nominal))
		       : fmaxf(correction, fminf(0.f, -1.f - nominal));
	}

	Vector update(const Vector &acceleration, const Vector &effectiveness, float dt,
		      float bandwidth_hz, float motor_tau, float limit, float slew_rate, bool usable, bool compensate)
	{
		if (!std::isfinite(dt) || dt <= 0.f || dt > 0.02f) {
			reset();
			return _correction;
		}

		const float model_gain = dt / (fmaxf(motor_tau, 0.001f) + dt);
		const float gain = dt / (1.f / (6.283185307f * fmaxf(bandwidth_hz, 0.1f)) + dt);
		limit = fminf(fmaxf(limit, 0.f), 0.3f);
		const float step = fmaxf(slew_rate, 0.f) * dt;

		for (unsigned i = 0; i < 3; ++i) {
			if (!std::isfinite(acceleration[i]) || !std::isfinite(effectiveness[i]) || effectiveness[i] <= 0.f
			    || !std::isfinite(_command[i])) {
				_model_command[i] = _disturbance[i] = _correction[i] = 0.f;
				continue;
			}

			_model_command[i] += model_gain * (_command[i] - _model_command[i]);

			if (!usable) {
				_disturbance[i] = _correction[i] = 0.f;
				continue;
			}

			// Convert the acceleration residual back to normalized torque units.
			const float residual = acceleration[i] / effectiveness[i] - _model_command[i];
			_disturbance[i] += gain * (residual - _disturbance[i]);
			_disturbance[i] = fminf(fmaxf(_disturbance[i], -limit), limit);
			const float target = compensate ? -_disturbance[i] : 0.f;
			_correction[i] += fminf(fmaxf(target - _correction[i], -step), step);
		}

		return _correction;
	}

private:

	Vector _model_command{};
	Vector _command{};
	Vector _disturbance{};
	Vector _correction{};
};

} // namespace collision_recovery
