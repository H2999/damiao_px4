// SPDX-License-Identifier: BSD-3-Clause
#include <gtest/gtest.h>
#include <limits>
#include "ImpactDetector.hpp"
#include "RateDisturbanceObserver.hpp"

using collision_recovery::ImpactDetector;
using collision_recovery::RateDisturbanceObserver;

TEST(ImpactDetection, DisabledAndNormalFlightDoNotTrigger)
{
	ImpactDetector detector;
	detector.update(1000000, 200.f, 40.f, true, false, 35.f, 30.f);
	EXPECT_EQ(detector.events(), 0u);
	EXPECT_FLOAT_EQ(detector.varianceScale(1000000, 25.f), 1.f);
	detector.update(1010000, 9.81f, 0.f, false, true, 35.f, 30.f);
	EXPECT_EQ(detector.state(), ImpactDetector::State::Idle);
}

TEST(ImpactDetection, ClippingTriggersWindowAndVarianceDecays)
{
	ImpactDetector detector;
	detector.update(1000000, 9.81f, 0.f, true, true, 35.f, 30.f);
	EXPECT_EQ(detector.events(), 1u);
	EXPECT_FLOAT_EQ(detector.varianceScale(1079999, 25.f), 25.f);
	EXPECT_EQ(detector.stateAt(1080000), ImpactDetector::State::Recovery);
	EXPECT_NEAR(detector.varianceScale(1330000, 25.f), 13.f, 1e-5f);
	EXPECT_FLOAT_EQ(detector.varianceScale(1580000, 25.f), 1.f);
	EXPECT_EQ(detector.stateAt(1580000), ImpactDetector::State::Idle);
}

TEST(ImpactDetection, RetriggerExtendsWindowWithoutCountingEverySample)
{
	ImpactDetector detector;
	detector.update(1000000, 40.f, 0.f, false, true, 35.f, 30.f);
	detector.update(1010000, 40.f, 0.f, false, true, 35.f, 30.f);
	EXPECT_EQ(detector.events(), 1u);
	EXPECT_EQ(detector.stateAt(1080000), ImpactDetector::State::Shock);
	detector.update(2000000, 9.81f, 31.f, false, true, 35.f, 30.f);
	EXPECT_EQ(detector.events(), 2u);
}

TEST(ImpactDetection, DuplicateAndReversedTimeAndNonfiniteInputs)
{
	ImpactDetector detector;
	detector.update(1000000, 9.81f, 0.f, false, true, 35.f, 30.f);
	detector.update(1000000, 100.f, 0.f, true, true, 35.f, 30.f);
	EXPECT_EQ(detector.events(), 0u);
	detector.update(500000, 100.f, 0.f, false, true, 35.f, 30.f);
	EXPECT_EQ(detector.events(), 1u);
	detector.update(2000000, std::numeric_limits<float>::quiet_NaN(), 0.f, false, true, 35.f, 30.f);
	EXPECT_EQ(detector.state(), ImpactDetector::State::Idle);
	EXPECT_FLOAT_EQ(detector.varianceScale(2000000, 1000.f), 1.f);
}

TEST(ImpactDob, ZeroEffectivenessDisablesUnidentifiedAxes)
{
	RateDisturbanceObserver dob;
	const auto output = dob.update({100.f, 100.f, 100.f}, {0.f, 0.f, 0.f}, .002f, 5.f, .04f, .1f, .5f, true, true);

	for (float value : output) { EXPECT_FLOAT_EQ(value, 0.f); }
}

TEST(ImpactDob, CorrectionUsesOnlyAvailableAuthorityAndPreservesSign)
{
	EXPECT_NEAR(RateDisturbanceObserver::boundedCorrection(.98f, .1f), .02f, 1e-6f);
	EXPECT_NEAR(RateDisturbanceObserver::boundedCorrection(-.98f, -.1f), -.02f, 1e-6f);
	EXPECT_FLOAT_EQ(RateDisturbanceObserver::boundedCorrection(1.5f, .1f), 0.f);
	EXPECT_FLOAT_EQ(RateDisturbanceObserver::boundedCorrection(-1.5f, -.1f), 0.f);
	EXPECT_FLOAT_EQ(RateDisturbanceObserver::boundedCorrection(1.5f, -.1f), -.1f);
	EXPECT_FLOAT_EQ(RateDisturbanceObserver::boundedCorrection(-1.5f, .1f), .1f);
}

TEST(ImpactDob, NonfiniteMeasurementCannotPoisonLaterUpdates)
{
	RateDisturbanceObserver dob;
	const float nan = std::numeric_limits<float>::quiet_NaN();
	EXPECT_FLOAT_EQ(dob.update({nan, 0.f, 0.f}, {100.f, 100.f, 100.f}, .002f,
				   5.f, .04f, .1f, .5f, true, true)[0], 0.f);
	const auto recovered = dob.update({5.f, 0.f, 0.f}, {100.f, 100.f, 100.f}, .002f,
					  5.f, .04f, .1f, .5f, true, true);
	EXPECT_TRUE(std::isfinite(recovered[0]));
	EXPECT_LT(recovered[0], 0.f);
}

TEST(ImpactDob, MonitorDoesNotApplyAndInvalidSensorInhibits)
{
	RateDisturbanceObserver dob;

	for (int i = 0; i < 1000; ++i) {
		const auto output = dob.update({5.f, 0.f, 0.f}, {100.f, 100.f, 100.f}, .002f, 5.f, .04f, .1f, .5f, true, false);
		EXPECT_FLOAT_EQ(output[0], 0.f);
	}

	EXPECT_NEAR(dob.disturbance()[0], .05f, 1e-5f);
	EXPECT_FLOAT_EQ(dob.update({5.f, 0.f, 0.f}, {100.f, 100.f, 100.f}, .002f, 5.f, .04f, .1f, .5f, false, true)[0], 0.f);
}

TEST(ImpactDob, IdentifiedPlantConvergesWithCorrectSignAndBoundedSlew)
{
	RateDisturbanceObserver dob;
	float motor_command = 0.f;
	float command = 0.f;
	float last_correction = 0.f;
	constexpr float dt = .002f;
	constexpr float effectiveness = 100.f;
	constexpr float external_accel = 5.f;
	float measured_accel = 0.f;

	for (int i = 0; i < 2000; ++i) {
		motor_command += dt / (.04f + dt) * (command - motor_command);
		measured_accel = effectiveness * motor_command + external_accel;
		const auto correction = dob.update({measured_accel, 0.f, 0.f}, {effectiveness, 100.f, 100.f},
						   dt, 5.f, .04f, .1f, .5f, true, true);
		EXPECT_LE(fabsf(correction[0] - last_correction), .5f * dt + 1e-7f);
		EXPECT_LE(fabsf(correction[0]), .1f);
		command = correction[0];
		dob.setCommand({command, 0.f, 0.f});
		last_correction = command;
	}

	EXPECT_NEAR(dob.disturbance()[0], .05f, 1e-4f);
	EXPECT_NEAR(command, -.05f, 1e-4f);
	EXPECT_NEAR(measured_accel, 0.f, 1e-3f);
}

TEST(ImpactDob, ShockDoesNotLearnImpulseAndRetainsActuatorModel)
{
	RateDisturbanceObserver dob;
	dob.setCommand({.2f, 0.f, 0.f});

	for (int i = 0; i < 500; ++i) {
		EXPECT_FLOAT_EQ(dob.update({10000.f, 0.f, 0.f}, {100.f, 100.f, 100.f}, .002f, 5.f, .04f, .1f, .5f, false, true)[0], 0.f);
	}

	// After measurement recovery, the observer should not mistake commanded torque for a disturbance.
	dob.update({20.f, 0.f, 0.f}, {100.f, 100.f, 100.f}, .002f, 5.f, .04f, .1f, .5f, true, true);
	EXPECT_NEAR(dob.disturbance()[0], 0.f, 1e-6f);
}

TEST(ImpactDob, CompensationReleasesSmoothlyAndLimitsLargeResidual)
{
	RateDisturbanceObserver dob;

	for (int i = 0; i < 1000; ++i) {
		dob.update({1000.f, 0.f, 0.f}, {100.f, 100.f, 100.f}, .002f, 5.f, .04f, .1f, .5f, true, true);
	}

	EXPECT_NEAR(dob.correction()[0], -.1f, 1e-6f);
	const float prior = dob.correction()[0];
	dob.update({0.f, 0.f, 0.f}, {100.f, 100.f, 100.f}, .002f, 5.f, .04f, .1f, .5f, true, false);
	EXPECT_NEAR(dob.correction()[0] - prior, .001f, 1e-6f);
	dob.update({0.f, 0.f, 0.f}, {100.f, 100.f, 100.f}, .1f, 5.f, .04f, .1f, .5f, true, true);
	EXPECT_FLOAT_EQ(dob.correction()[0], 0.f);
}
