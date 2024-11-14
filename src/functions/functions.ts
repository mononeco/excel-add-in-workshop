/* global clearInterval, CustomFunctions, setInterval */

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * generate zebra pattern.
 * @customfunction
 * @param invocation Custom function handler
 */
export function zebra(invocation: CustomFunctions.StreamingInvocation<number[][]>): void {
  const gs = new GrayScotModel(0.04, 0.06, 100);

  const timer = setInterval(() => {
    gs.update();
    invocation.setResult(gs.materialU);
  }, 10);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

class GrayScotModel {
  feed: number;
  kill: number;
  materialU: number[][];
  materialV: number[][];
  spaceGridSize: number;

  readonly Dx = 0.01;
  readonly Dt = 1;
  readonly Du = 2e-5;
  readonly Dv = 1e-5;

  constructor(feed: number, kill: number, spaceGridSize: number) {
    this.feed = feed;
    this.kill = kill;
    this.spaceGridSize = spaceGridSize;
    this.materialU = Array.from(new Array(spaceGridSize), () => new Array(spaceGridSize).fill(1));
    this.materialV = Array.from(new Array(spaceGridSize), () => new Array(spaceGridSize).fill(0));

    const fromRange = Math.floor(spaceGridSize / 2) - Math.floor(spaceGridSize / 20);
    const toRange = Math.floor(spaceGridSize / 2) + Math.floor(spaceGridSize / 20);

    for (let i = fromRange; i < toRange; i++) {
      for (let j = fromRange; j < toRange; j++) {
        this.materialU[i][j] = 0.5;
        this.materialV[i][j] = 0.25;
      }
    }

    for (let i = 0; i < spaceGridSize; i++) {
      for (let j = 0; j < spaceGridSize; j++) {
        this.materialU[i][j] += Math.random() * 0.1;
        this.materialV[i][j] += Math.random() * 0.1;
      }
    }
  }

  update() {
    const laplacianU: number[][] = Array.from(new Array(this.spaceGridSize), () => new Array(this.spaceGridSize));
    const laplacianV: number[][] = Array.from(new Array(this.spaceGridSize), () => new Array(this.spaceGridSize));

    const firstIndex = 0;
    const lastIndex = this.spaceGridSize - 1;

    for (let i = 0; i < this.spaceGridSize; i++) {
      for (let j = 0; j < this.spaceGridSize; j++) {
        const rollDownIndex = i === 0 ? lastIndex : i - 1;
        const rollUpIndex = i === lastIndex ? firstIndex : i + 1;
        const rollRightIndex = j === 0 ? lastIndex : j - 1;
        const rollLeftIndex = j === lastIndex ? firstIndex : j + 1;

        laplacianU[i][j] =
          (this.materialU[rollDownIndex][j] +
            this.materialU[rollUpIndex][j] +
            this.materialU[i][rollRightIndex] +
            this.materialU[i][rollLeftIndex] -
            4 * this.materialU[i][j]) /
          (this.Dx * this.Dx);
        laplacianV[i][j] =
          (this.materialV[rollDownIndex][j] +
            this.materialV[rollUpIndex][j] +
            this.materialV[i][rollRightIndex] +
            this.materialV[i][rollLeftIndex] -
            4 * this.materialV[i][j]) /
          (this.Dx * this.Dx);
      }
    }

    for (let i = 0; i < this.spaceGridSize; i++) {
      for (let j = 0; j < this.spaceGridSize; j++) {
        const dudt =
          this.Du * laplacianU[i][j] -
          this.materialU[i][j] * this.materialV[i][j] * this.materialV[i][j] +
          this.feed * (1.0 - this.materialU[i][j]);
        const dvdt =
          this.Dv * laplacianV[i][j] +
          this.materialU[i][j] * this.materialV[i][j] * this.materialV[i][j] -
          (this.feed + this.kill) * this.materialV[i][j];
        this.materialU[i][j] += this.Dt * dudt;
        this.materialV[i][j] += this.Dt * dvdt;
      }
    }
  }
}
