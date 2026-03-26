export declare class Length {
    readonly value: number;
    readonly type?: string;
    constructor(value: number, type?: string);
    static parse(text: string): Length;
    static from(val: any): Length;
    add(length: Length): Length;
    mul(val: number): Length;
    valueOf(): number;
    toString(): string;
}
