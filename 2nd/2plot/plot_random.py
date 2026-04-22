import os
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np

matplotlib.rcParams["font.family"] = "Malgun Gothic"
matplotlib.rcParams["axes.unicode_minus"] = False  # 마이너스 기호 깨짐 방지


def plot_random_histogram(
    n: int = 100,
    mean: float = 50.0,
    std: float = 15.0,
    bins: int = 15,
    output_path: str = "histogram.pdf",
    seed: int | None = None,
) -> str:
    """
    정규분포 랜덤 숫자 n개를 생성하고 히스토그램을 PDF로 저장.
    반환값: 저장된 PDF 절대 경로
    """
    rng = np.random.default_rng(seed)
    data = rng.normal(loc=mean, scale=std, size=n)

    fig, ax = plt.subplots(figsize=(10, 7))

    # 히스토그램
    counts, edges, patches = ax.hist(
        data, bins=bins,
        color="#4C72B0", edgecolor="white", linewidth=0.8, alpha=0.85,
        label=f"데이터 ({n}개)",
    )

    # 평균선
    mu = data.mean()
    sigma = data.std()
    ax.axvline(mu, color="#C44E52", linestyle="--", linewidth=2,
               label=f"평균 (μ) = {mu:.2f}")

    # ±1σ
    ax.axvline(mu - sigma, color="#DD8452", linestyle=":", linewidth=1.8,
               label=f"±1σ  (σ = {sigma:.2f})")
    ax.axvline(mu + sigma, color="#DD8452", linestyle=":", linewidth=1.8)

    # ±2σ 범위 배경 강조
    ax.axvspan(mu - sigma, mu + sigma, alpha=0.08, color="#DD8452")

    # 통계 요약 텍스트 박스
    stats_text = (
        f"n  = {n}\n"
        f"μ  = {mu:.3f}\n"
        f"σ  = {sigma:.3f}\n"
        f"min = {data.min():.3f}\n"
        f"max = {data.max():.3f}"
    )
    ax.text(
        0.97, 0.97, stats_text,
        transform=ax.transAxes,
        fontsize=10, verticalalignment="top", horizontalalignment="right",
        bbox=dict(boxstyle="round,pad=0.5", facecolor="white",
                  edgecolor="#BBBBBB", alpha=0.9),
        fontfamily="monospace",
    )

    # 레이블 · 제목
    ax.set_title(f"정규분포 랜덤 숫자 {n}개 — 히스토그램", fontsize=15, pad=16)
    ax.set_xlabel("값", fontsize=12, labelpad=8)
    ax.set_ylabel("빈도 (개수)", fontsize=12, labelpad=8)
    ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    ax.legend(fontsize=11)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.spines[["top", "right"]].set_visible(False)

    plt.tight_layout()

    abs_path = os.path.abspath(output_path)
    plt.savefig(abs_path, format="pdf", bbox_inches="tight")
    plt.close(fig)
    return abs_path


if __name__ == "__main__":
    out = "E:/venture/proposal/JBLab/2nd/2plot/histogram.pdf"
    path = plot_random_histogram(seed=42, output_path=out)
    print(f"PDF 저장: {path}")
