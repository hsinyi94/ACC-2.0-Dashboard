from weekly_gms import build_weekly_gms
result = build_weekly_gms(force_all=True)
print(f"完成: {len(result.weeks)} 週, 最新 wk{result.weeks[-1]}")
