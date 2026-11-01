// SPDX-License-Identifier: BSD-3-Clause
#include <gtest/gtest.h>
#include "EKF/ekf.h"
#include "EKF/imu_down_sampler/imu_down_sampler.hpp"

namespace
{
imuSample sample(uint64_t time_us)
{
	imuSample imu{};
	imu.time_us = time_us;
	imu.delta_ang_dt = .01f;
	imu.delta_vel_dt = .01f;
	imu.delta_vel = Vector3f(0.f, 0.f, -9.80665f * .01f);
	return imu;
}

void feed(Ekf &ekf, const imuSample &imu)
{
	ekf.setIMUData(imu);
	ekf.update();
}

void initialise(Ekf &ekf)
{
	ekf.init(0);
	ekf.set_in_air_status(false);

	for (uint64_t time = 10000; time <= 2000000; time += 10000) { feed(ekf, sample(time)); }
}
}

TEST(EkfImpact, GyroClippingSurvivesDownsamplingAndResets)
{
	int32_t target_dt = 10000;
	ImuDownSampler sampler(target_dt);
	imuSample imu = sample(10000);
	imu.delta_ang_dt = imu.delta_vel_dt = .001f;
	imu.delta_vel = Vector3f(0.f, 0.f, -.00980665f);
	bool ready = false;

	for (unsigned i = 0; i < 30 && !ready; ++i) {
		imu.time_us += 1000;
		imu.delta_ang_clipping[1] = (i == 0);
		ready = sampler.update(imu);
	}

	ASSERT_TRUE(ready);
	const auto clipped = sampler.getDownSampledImuAndTriggerReset();
	EXPECT_TRUE(clipped.delta_ang_clipping[1]);
	EXPECT_FALSE(clipped.delta_ang_clipping[0]);
	ready = false;
	imu.delta_ang_clipping[1] = false;

	for (unsigned i = 0; i < 30 && !ready; ++i) {
		imu.time_us += 1000;
		ready = sampler.update(imu);
	}

	ASSERT_TRUE(ready);
	EXPECT_FALSE(sampler.getDownSampledImuAndTriggerReset().delta_ang_clipping[1]);
}

TEST(EkfImpact, DisabledByDefaultAndGroundShockDoesNotAdapt)
{
	Ekf ekf;
	initialise(ekf);
	EXPECT_EQ(ekf.getParamHandle()->ekf2_imp_en, 0);
	ekf.getParamHandle()->ekf2_imp_en = 1;

	for (uint64_t time = 2010000; time < 2300000; time += 10000) {
		auto imu = sample(time);
		imu.delta_ang_clipping[0] = true;
		feed(ekf, imu);
	}

	EXPECT_FLOAT_EQ(ekf.getImpactVarianceScale(), 1.f);
	EXPECT_EQ(ekf.getImpactDetector().events(), 0u);
}

TEST(EkfImpact, AirborneClippingInflatesCovarianceAndReturnsToNominal)
{
	Ekf baseline;
	Ekf adaptive;
	initialise(baseline);
	initialise(adaptive);
	adaptive.getParamHandle()->ekf2_imp_en = 1;
	baseline.set_in_air_status(true);
	adaptive.set_in_air_status(true);

	for (uint64_t time = 2010000; time <= 2300000; time += 10000) {
		auto imu = sample(time);
		imu.delta_ang_clipping[0] = true;
		feed(baseline, imu);
		feed(adaptive, imu);
	}

	EXPECT_FLOAT_EQ(baseline.getImpactVarianceScale(), 1.f);
	EXPECT_FLOAT_EQ(adaptive.getImpactVarianceScale(), 25.f);
	EXPECT_GT(adaptive.getTiltVariance(), baseline.getTiltVariance());
	EXPECT_TRUE(adaptive.getVelocity().isAllFinite());
	EXPECT_NEAR(adaptive.getVelocity()(0), baseline.getVelocity()(0), 1e-5f);

	for (uint64_t time = 2310000; time <= 3300000; time += 10000) { feed(adaptive, sample(time)); }

	EXPECT_FLOAT_EQ(adaptive.getImpactVarianceScale(), 1.f);
	EXPECT_TRUE(adaptive.getVelocityVariance().isAllFinite());
}

TEST(EkfImpact, SpecificForceImpulseIsNotRemovedFromPropagation)
{
	Ekf ekf;
	initialise(ekf);
	ekf.getParamHandle()->ekf2_imp_en = 1;
	ekf.set_in_air_status(true);

	for (uint64_t time = 2010000; time <= 2200000; time += 10000) {
		auto imu = sample(time);
		imu.delta_vel(0) = 40.f * imu.delta_vel_dt;
		feed(ekf, imu);
	}

	EXPECT_GT(ekf.getImpactVarianceScale(), 1.f);
	EXPECT_GT(ekf.getVelocity()(0), 1.f);
}
